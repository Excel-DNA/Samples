# Excel-DNA Testing Helper

_**Note: The Excel-DNA Testing Helper is in an early preview stage, and the design and functionality might change a lot in future. If you try it and have any feedback at all, positive or 'constructive' please let me know at govert@icon.co.za. If you already practice automated testing extensively I am particularly eager to hear of any suggestions you might have.**_

[![Excel-DNA Testing Helper - Preview](http://img.youtube.com/vi/rUP15u9Z0ik/0.jpg)](http://www.youtube.com/watch?v=rUP15u9Z0ik "Excel-DNA Testing Helper - Preview")

`ExcelDna.Testing` is a NuGet package and library that lets you develop automatic tests for Excel models and add-ins, including add-ins developed with Excel-DNA and VBA. Test code is written in C# or Visual Basic and is hosted by the popular [xUnit](https://xunit.net/) test framework, allowing automated tests to run from Visual Stuio or other standard test runners.

Tests developed with the testing helper will run with the add-in loaded inside an Excel instance, and so allow you to test the interaction between an add-in and a real instance of Excel. This type of 'integration testing' can augment 'unit testing' where individual library features are tested in isolation. It is often in the interaction with Excel where the problematic aspects of an add-in are revealed, and developing automated testing for this environment has been difficult.

The testing helper allows flexibility and power in designing automated Excel tests:
* The test code can either run in a separate process that drives Excel through the COM object model, or can be loaded inside the Excel process itself, allowing use of both the COM object model and the C API from the test code.
* Functions, macros and even ribbon commands can be tested.
* Test projects can include pre-populated test workbooks containing spreadsheet models to test or test data.

Running automated tests against Excel does introduce complications:
* Testing requires a copy of Excel to be installed on the machine where the tests are run, so don't work well as automated test for 'continuous integration' environments.
* Test outcomes can depend on the exact version of Excel the is used. This is both an advantage in identifying some 
* Integration tests with Excel can be quite slow to run compared to direct unit testing of functions.

This tutorial will introduce the Excel-DNA testing helper, and show you how to create a test project for your Excel model or add-in.

## Background and prerequisites

* **Visual Studio and Excel** - 
To use the testing helper you should already have Visual Studio 2019 and Excel (any version) installed.
The example will mostly use C#, but Visual Basic is fully supported can also be used for creating your test project.

* **xUnit** - 
[xUnit](https://xunit.net/) is a unit testing tool for the .NET Framework. The required xUnit libraries and runner will automatically installed with the `ExcelDna.Testing` package.

If you are not familiar with unit test frameworks, or with xUnit in particular, you might want to look at or work through the XUnit Getting Started instructions for 
[Using .NET Framework with Visual Studio](https://xunit.net/docs/getting-started/netfx/visual-studio).

## Create a test project
To start a new test project:
* create a new  'Class Library (.NET Framework)' project (using C# or Visual Basic) and
* install the `ExcelDna.Testing` package from the NuGet package manager (currently a pre-release package, so check the relevant checkbox or add the `-Pre` flag to the NuGet command line).

After installing the `ExcelDna.Testing` package, the project will have the xUnit framework and Visual Studio runner for xUnit installed, so no additional packages are needed.

## Testing examples

### *ExcelTest* - A simple Excel test

This is a simple test the exercises Excel to 

```c#
using System;
using Xunit;
using Microsoft.Office.Interop.Excel;
using ExcelDna.Testing;

// This attribute MUST be present somewhere in the test project to connect xUnit to the ExcelDna.Testing framework.
[assembly: TestFramework("Xunit.ExcelTestFramework", "ExcelDna.Testing")]

namespace ExcelTest
{
    public class CalculationTests : IDisposable
    {
        Workbook _testWorkbook;

        public CalculationTests()
        {
            // Get hold of the Excel Application object and create a workbook
            _testWorkbook = Util.Application.Workbooks.Add();
        }

        public void Dispose()
        {
            // Clean up our workbook without saving changes
            _testWorkbook.Close(SaveChanges: false);
        }

        [ExcelFact]
        public void NumbersAddCorrectly()
        {
            // We'll just do our test on the first sheet
            var ws = _testWorkbook.Sheets[1];

            // Write two numbers to the active sheet, and a formula that adds them, together
            ws.Range["A1"].Value = 2.0;
            ws.Range["A2"].Value = 3.0;
            ws.Range["A3"].Formula = "= A1 + A2";

            // Read back the value from the cell with the formula
            var result = ws.Range["A3"].Value;

            // Check that we have the expected result
            Assert.Equal(5.0, result);
        }
    }
}
```

To run the tests in Visual Studio, open the Test Explorer tool window, check that the test is correctly discovered, and press Run.

#### Discussion

Some notable aspects of the above code snippet:
* The `Xunit.ExcelTestFramework` is configured through the `Xunit.TestFramework` assembly-scope attribute.
* Tests are public instance methods marked by an `[ExcelDna.Testing.ExcelFact]` attribute.
* Test code can access Excel `Application` object with a call to `ExcelDna.Testing.Util.Application`. This will refer to the correct Excel root COM object, whether the test code is running in-process or out-of-process (see below).
* We use the class constructor and `IDispose` interface to set up and tear down some 

### *AddInTest* - Testing an Excel-DNA add-in

For this test project we create a simple Excel-DNA add-in with a single UDF, and then implement a test project that exercises the add-in function inside Excel.

```c#
using System;
using Xunit;
using ExcelDna.Testing;
using Microsoft.Office.Interop.Excel;

// This attribute MUST be present somewhere in the test project to connect xUnit to the ExcelDna.Testing framework.
// It could also be placed in the Properties\AssemblyInfo.cs file.
[assembly:Xunit.TestFramework("Xunit.ExcelTestFramework", "ExcelDna.Testing")]

namespace Sample.Test
{
    // The path give here is relative to the output directory of the test project.
    // Setting an AddIn options here will request the test runner to load this add-in into Excel before the tests start.
    // The name here excludes the ".xll" or "64.xll" suffix. The test runner will choose according to the Excel bitness where it runs.
    [ExcelTestSettings(AddIn = @"..\..\..\Sample\bin\Debug\Sample-AddIn")]
    public class ExcelTests : IDisposable
    {
        // This workbook will be available to all tests in the class
        Workbook _testWorkbook;

        // The test class constructor will configure the required environment for the tests in the class.
        // In this case it creates a new Workbook that will be shared by the tests
        public ExcelTests()
        {
            var app = Util.Application;
            _testWorkbook = app.Workbooks.Add();
        }

        // Clean-up for the class is in the IDisposable.Dispose implementation
        public void Dispose()
        {
            _testWorkbook.Close(SaveChanges: false);
        }

        // This test just interacts with Excel
        [ExcelFact]
        public void ExcelCanAddNumbers()
        {
            var ws = _testWorkbook.Sheets[1];

            ws.Range["A1"].Value = 2.0;
            ws.Range["A2"].Value = 3.0;
            ws.Range["A3"].Formula = "=A1 + A2";

            var result = ws.Range["A3"].Value;

            Assert.Equal(5.0, result);
        }

        // This test depends on the AddIn value set in the class's ExcelTestSettings attributes
        // With the Sample-AddIn loaded, the function should work correctly.
        [ExcelFact]
        public void AddInCanAddNumbers()
        {
            var ws = _testWorkbook.Sheets[1];

            ws.Range["A1"].Value = 2.0;
            ws.Range["A2"].Value = 3.0;
            ws.Range["A3"].Formula = "=AddThem(A1, A2)";

            var result = ws.Range["A3"].Value;

            Assert.Equal(5.0, result);
        }

        // Before this test is run, a pre-created workbook will be loaded
        // It has been added to the test project and configured to always be copied to the output directory
        [ExcelFact(Workbook = "TestBook.xlsx")]
        public void WorkbookCheckIsOK()
        {
            // Get the pre-loaded workbook using the Util.Workbook property
            var wb = Util.Workbook;
            var ws = wb.Sheets["Check"];
            Util.Application.CalculateFull();

            var result = ws.Range["A1"].Value;

            Assert.Equal("OK", result);
        }
    }
}

```

#### Discussion

This snippet is from the accompanying sample solution, which also contains an Excel-DNA add-in project (creating the `Sample-AddIn` add-in referred to in the `ExcelTestSettings` attribute.
The add-in contains a single function called `AddThem` which is given a first test in the sample test project.

Note that the sample test project does not reference the add-in project - all interaction is through Excel. This ensures that we are truly testing the behaviour of the add-in when running in Excel.

## Solution layout suggestion

For supporting both isolated unit testing and Excel-based integration testing, one possible solution layout is as follows:

* **MyLibrary** - contains the core functionality, e.g. calculations or external data access methods. Does not reference Excel-DNA or Excel.
* **MyLibrary.Test** - unit test project for the functionality in `MyLibrary`, using the standard `xunit` and `xunit.runner.visualstudio` packages as described in the [xUnit documentation](https://xunit.net/docs/getting-started/netfx/visual-studio).
* **MyAddIn** - Excel AddIn to integrate the functionality from `MyLibrary` into Excel, using the `ExcelDna.AddIn` package. Functions declared here contain Excel-specific attributes and information, and deal with the Excel data types and error values if needed before calling into 'Library' methods.
* **MyAddIn.Test** - integration testing project for 'MyAddIn', using the `ExcelDna.Testing` package. Does not reference the `MyAddIn` or `MyLibrary` projects, just interacts with them through the Excel tests.

## Reference

### Test helper classes 

These types are declared in the `ExcelDna.Testing` assembly.

* *`Xunit.ExcelTestFramework`* - this is the xUnit integration class, and should be indicated in the assembly-scope TestFramework attribute inside the test project:
```c#
[assembly:TestFramework("Xunit.ExcelTestFramework", "ExcelDna.Testing")]
```

* *`ExcelDna.Testing.ExcelFactAttribute`* - this is the method-scope attribute to indicate that a method implements a test.
```c#
      [ExcelFact]
      public void NumbersAddCorrectly() { ... }
```
* *`ExcelDna.Testing.ExcelTestSettingsAttribute`* - this is a class-scope attribute that configures settings for all tests in a class.
```c#
      [ExcelTestSettings(OutOfProcess=true)]
      public class CalculationTests { ... }
```

* *`ExcelDna.Testing.Util`* - provides access to the root Excel Application object, any pre-loaded Workbook and the directory where the test assembly is located.

### In-process vs out-of-process test running
The test methods can execute in two environments:

* Inside the Excel process (the default) - The test runner will load a helper add-in (called ExcelAgent) into the Excel process, and ExcelAgent in turn will load the test library. Test will then run inside the Excel process, which improves performance and gives access to the the Excel C API - `XlCall.Excel(...)` - from the test code.

* Out-of-Process  - There is also an option to run tests out-of-process. This is indicated by setting the `OutOfProcess` property of the `ExcelFact` or `ExcelSettings` attribute on the methopd or class respectively. In this case the test assembly will run inside the xUnit test runner and communicate with Excel via the stardard cross-process COM interop. In this approach there is no additional test agent loaded into the Excel process.

### Error values - COM vs C API

One of the motivations for doing integration testing of an add-in in Excel is to ensure the behaviour of a function when receiving various unexpected values from Excel is correct. In particular, a function running in Excel might receive input values like 'Empty', 'Missing' or some Excel-specific error value like '#VALUE'. Depending on how the test code is reading values from Excel, these additional data types would be represented in different ways.
It is not discussed here, but when running inside the Excel process, the C API (XlCall.Excel) can often improve the performance of Excel sheet interactions, and simplify dealing with the various Excel error values.

### Thanks

Thank you very much to Sergey Vlasov for developing the `ExcelDna.Testing` framework.

