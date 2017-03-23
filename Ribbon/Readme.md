# Ribbon Sample

This sample shows how to add a ribbon UI extension with the Excel-DNA add-in, and how to use the Excel COM object model from C# to write some information to a workbook.

## Initial setup

The initial setup will create a new add-in with a simple test function (it's a useful indicator to show that the add-in is loaded into the Excel session).

1. Create new Class Library project.
2. Install `ExcelDna.AddIn` package.
3. Add a small test function:

```cs
namespace Ribbon
{
    public static class Functions
    {
        public static string dnaRibbonTest()
        {
            return "Hello from the Ribbon Sample!";
        }
    }
}
```

4. Press F5 to load in Excel, and then test `=dnaRibbonTest()` in a cell.

## Add the ribbon controller

Next we add a class to implement the ribbon UI extension, with a simple button.

1. Add a reference to the `System.Windows.Forms` assembly (we'll use that for showing our messages).

2. Add a new class for the ribbon controller (maybe `RibbonController.cs`), with this code for a button and handler:

```cs
using System.Runtime.InteropServices;
using System.Windows.Forms;
using ExcelDna.Integration.CustomUI;

namespace Ribbon
{
    [ComVisible(true)]
    public class RibbonController : ExcelRibbon
    {
        public override string GetCustomUI(string RibbonID)
        {
            return @"
      <customUI xmlns='http://schemas.microsoft.com/office/2006/01/customui'>
      <ribbon>
        <tabs>
          <tab id='tab1' label='My Tab'>
            <group id='group1' label='My Group'>
              <button id='button1' label='My Button' onAction='OnButtonPressed'/>
            </group >
          </tab>
        </tabs>
      </ribbon>
    </customUI>";
        }

        public void OnButtonPressed(IRibbonControl control)
        {
            MessageBox.Show("Hello from control " + control.Id);
        }
    }
}
```

3. Press F5 to load and test.


#### Notes

* The ribbon class derives from the `ExcelDna.Integration.CustomUI.ExcelRibbon` base class. This is how Excel-DNA itentifies the class a defining a ribbon controller.

* The ribbon class must be 'COM visible'. Either the class must be marked as `[ComVisible(true)]` (the default class library template in Visual Studio markes the assembly as `[assembly:ComVisible(false)]`).

* The xml namespace is important. Excel 2007 introduced the ribbon, and support only the original xml namespace as shown in this example - `xmlns='http://schemas.microsoft.com/office/2006/01/customui'`. Further enhancements to the ribbon was made in Excel 2010, including using the ribbon for worksheet context menus and adding the backstage area. To indicate the extended Excel 2010 xml schema, this version and later supports an update namespace - `xmlns='http://schemas.microsoft.com/office/2009/07/customui'`.

* The Office applications have a debugging setting to assist in finding any errors in the ribbon xml, which would prevent the ribbon from loading. In Excel 2013, this setting can be found under `File -> Options -> Advanced`, then under `General` find 'Show add-in user interface errors'. Note that this setting applied to all installed Office applications, and can reveal unexpected errors that are present in other add-ins too.

* There are different options for providing the ribbon xml. In this sample it is embedded as a string in the code and returned from the `ExcelRibbon.GetCustomUI` overload. Excel-DNA also supports placing the xml inside the .dna add-in configuration file (this is where the base class implementation of `GetCustomUI` looks for it). The ribbon xml can also be put in an assembly resource (either as a string or from a separate file) and extracted at runtime with some extra code in `GetCustomUI`.

* The callback methods, like `OnButtonPressed` in the example, are found by Excel using the COM `IDispatch` interface that is implicitly implemented by the COM visible .NET class.

* Behind the scenes, Excel-DNA registers and loads a COM helper add-in that provides the ribbon support. This COM helper add-in should load even if the user does not have administrator rights, but it might be blocked by some Excel-specific security settings.

* Errors in the ribbon methods can cause Excel to mark the ribbon COM helper add-in as a 'Disabled Add-in'. This will reflect in the 'Disabled Add-ins' list under `File-> Options -> Add-Ins` under the `Manage` dropdown.

#### Ribbon xml and callback documentation

Excel-DNA is responsible for loading the ribbon helper add-in, but is not otherwise involved in the ribbon extension. This means that the custom UI xml schema, and the signatures for the callback methods are exactly as documented by Microsoft. The best documentation for these aspects can be found in the three-part series on 'Customizing the 2007 Office Fluent Ribbon for Developers':

* [Part 1 - Overview](https://msdn.microsoft.com/en-us/library/aa338202.aspx)
* [Part 2 - Controls and callback reference](https://msdn.microsoft.com/en-us/library/aa338199.aspx)
* [Part 3 - Frequently asked questions, including C# and VB.NET callback signatures](https://msdn.microsoft.com/en-us/library/aa722523.aspx)

Information related to the Excel 2010 extensions to the ribbon can be found here:

* [Customizing Context Menus in Office 2010](https://msdn.microsoft.com/en-us/library/office/ee691832.aspx)
* [Customizing the Office 2010 Backstage View](https://msdn.microsoft.com/en-us/library/office/ee815851.aspx)
* [Ribbon Extensibility in Office 2010: Tab Activation and Auto-Scaling](https://msdn.microsoft.com/en-us/library/office/ee691834.aspx)

Creating Dynamic Ribbon Customizations
* [Part 1](https://msdn.microsoft.com/en-us/library/dd548010%28v=office.12%29.aspx)
* [Part 2](https://msdn.microsoft.com/en-us/library/dd548011%28v=office.12%29.aspx)

Other ribbon-related resources:

* [Ron de Bruin's Excel Tips](http://www.rondebruin.nl/win/s2/win003.htm)
* [Andy Pope's RibbonX Visual Designer](http://www.andypope.info/vba/ribboneditor.htm)

## Add access to the Excel COM object model

In this step we add access to the Excel COM object model, to show how C# code can use the familiar object model to manipulate Excel.

1. Add a reference to the Primary Interop Assembly (PIA) for Excel. The easiest way to do this is to install the `ExcelDna.Interop` NuGet package. This will install the interop assemblies that correspond to Excel 2010, so are suitable for add-ins that support Excel 2010 and later.

2. Add a class for our data writer:

```cs
using System;
using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;

namespace Ribbon
{
    public class DataWriter
    {
        public static void WriteData()
        {
            Application xlApp = (Application)ExcelDnaUtil.Application;

            Workbook wb = xlApp.ActiveWorkbook;
            if (wb == null)
                return;

            Worksheet ws = wb.Worksheets.Add(Type: XlSheetType.xlWorksheet);
            ws.Range["A1"].Value = "Date";
            ws.Range["B1"].Value = "Value";

            Range headerRow = ws.Range["A1", "B1"];
            headerRow.Font.Size = 12;
            headerRow.Font.Bold = true;

            // Generally it's faster to write an array to a range
            var values = new object[100, 2];
            var startDate = new DateTime(2007, 1, 1);
            var rand = new Random();
            for (int i = 0; i < 100; i++)
            {
                values[i, 0] = startDate.AddDays(i);
                values[i, 1] = rand.NextDouble();
            }

            ws.Range["A2"].Resize[100, 2].Value = values;
            ws.Columns["A:A"].EntireColumn.AutoFit();

            // Add a chart
            Range dataRange= ws.Range["A1:B101"];
            dataRange.Select();
            ws.Shapes.AddChart(XlChartType.xlLineMarkers).Select();
            xlApp.ActiveChart.SetSourceData(Source: dataRange);
        }
    }
}
```

3. Update the ribbon handler to call our data writer:

```cs
        public void OnButtonPressed(IRibbonControl control)
        {
            MessageBox.Show("Hello from control " + control.Id);
            DataWriter.WriteData();
        }
```

4. Press F5 to load and press the ribbon button to run the `WriteData` code.

#### Getting the root `Application` object

A key call in the above code is to retrieve the root `Application` object that matches the Excel instance that is hosting the add-in. We call `ExcelDnaUtil.Application`, which returns an object that is always the correct `Application` COM object. Code that attempts to get the `Application` object in other ways, e.g. by calling `new Application()` might work in some cases, but there is a danger that the Excel instance returned is not that instance hosting the add-in.  

Once the root `Application` object is retrieved, the object model is accessed normally as it would be from VBA.

* Don't confuse the types `Microsoft.Office.Interop.Excel.Application` that we use here with the WinForms type `System.Windows.Forms.Application`. You might use a namespace alias to distinguish these in your code.

#### Interop assembly versions

Each version of Excel adds some extensions to the object model (and rarely, but sometimes removes some parts). The changes might be entire classes and interfaces, methods on an interface or parameters or a method. Most add-ins are expected to run on different Excel versions, so some care is needed to make sure only object models features available on all the target versions are used.

The simplest approach is to pick a minimum Excel version to support, and use the COM object model definitions (PIA asemblies) from that version. Such code will work against the chosen version, and any any other version (newer or older) that implements the same parts of the object model. Since most Excel versions only add to the object model, this means that add-in will work correctly with newer versions too. This is similar to developing a VBA extension on Excel 2010, which might then fail on older versions if the VBA code uses methods not available on the running version.

In this example we've installed the 'ExcelDna.Interop' NuGet package, which includes the interop assemblies for Excel 2010. This means features added in Excel 2013 and later will not be shown in the object model IntelliSense, ensuring that the add-in only uses features available on the minimum version. 

#### Correct COM / .NET interop usage

There is a lot of misinformation on the web about using the COM object model from .NET.

* To ensure that the Excel process always correctly exits, Excel add-ins should *only call the Excel COM object model from the main Excel thread, in a macro of callback context*. Never attempt to access the COM object model from multiple threads - since the Excel COM object model is single-threaded (technically a Single-Threaded Apartment) there can be no performance benefit in trying to access Excel from multiple threads.

* An Excel add-in should never call `Marshal.ReleaseComObject(...)` or `Marshal.FinalReleaseComObject(...)` when doing Excel interop. It is a confusing anti-pattern, but any information about this, including from Microsoft, that indicates one should manually release COM references from .NET is incorrect. The .NET runtime and garbage collector correctly keep track of and clean up COM references.

* Any guidance that mentions 'double-dots' is misleading. Sometimes this indicates that expressions calling into the object model should not chain object model access, i.e. to avoid code like `myWorkbook.Sheets[1].Range["A1"].Value = "abc". Such code is fine - just ignore any 'two dots' guidance.

* I've posted some more details on these issues (in the context of automating Excel from another application) in a [Stack Overflow answer](http://stackoverflow.com/a/38111294/44264).

## Further ribbon topics

These are some more aspects of the ribbon extensions and COM object model, not yet dealt with:

* Updating the ribbon, e.g. to trigger a `getEnabled` callback - the `onLoad` callback must be implemented to capture the ribbon interface during loading.

* Adding images to the ribbon. Note the the `IPictureDisp` interface does not have to be used - any `Bitmap` type can be returned from the `getImage` callbacks. Excel-DNA has some support for packing image files into the .xll.

* Using COM object model events.

* Transitioning to the main thread (or a macro context) from another thread. Excel-DNA has a helper method `ExcelAsyncUtil.QueueAsMacro` that can be called from another thread or a timer event, to transition to a context where the object model can be reliably used.
