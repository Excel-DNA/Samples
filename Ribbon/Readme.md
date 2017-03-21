# Ribbon Sample

This sample shows how to add a ribbon UI extension with the Excel-DNA add-in, and how to use the Excel COM object model from C# to write some information to a workbook.

## Initial setup

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


### Notes

* The ribbon class derives from the `ExcelDna.Integration.CustomUI.ExcelRibbon` base class. This is how Excel-DNA itentifies the class a defining a ribbon controller.

* The ribbon class must be 'COM visible'. Either the class must be marked as `[ComVisible(true)]` (the default class library template in Visual Studio markes the assembly as `[assembly:ComVisible(false)]`).

* The xml namespace is important. Excel 2007 introduced the ribbon, and support only the original xml namespace as shown in this example - `xmlns='http://schemas.microsoft.com/office/2006/01/customui'`. Further enhancements to the ribbon was made in Excel 2010, including using the ribbon for worksheet context menus and adding the backstage area. To indicate the extended Excel 2010 xml schema, this version and later supports an update namespace - `xmlns='http://schemas.microsoft.com/office/2009/07/customui'`.

* The Office applications have a debugging setting to assist in finding any errors in the ribbon xml, which would prevent the ribbon from loading. In Excel 2013, this setting can be found under `File -> Options -> Advanced`, then under `General` find 'Show add-in user interface errors'. Note that this setting applied to all installed Office applications, and can reveal unexpected errors that are present in other add-ins too.

* There are different options for providing the ribbon xml. In this sample it is embedded as a string in the code and returned from the `ExcelRibbon.GetCustomUI` overload. Excel-DNA also supports placing the xml inside the .dna add-in configuration file (this is where the base class implementation of `GetCustomUI` looks for it). The ribbon xml can also be put in an assembly resource (either as a string or from a separate file) and extracted at runtime with some extra code in `GetCustomUI`.

* The callback methods, like `OnButtonPressed` in the example, are found by Excel using the COM `IDispatch` interface that is implicitly implemented by the COM visible .NET class.

* Behind the scenes, Excel-DNA registers and loads a COM helper add-in that provides the ribbon support. This COM helper add-in should load even if the user does not have administrator rights, but it might be blocked by some Excel-specific security settings.

* Errors in the ribbon methods can cause Excel to mark the ribbon COM helper add-in as a 'Disabled Add-in'. This will reflect in the 'Disabled Add-ins' list under `File-> Options -> Add-Ins` under the `Manage` dropdown.

### Ribbon xml and callback documentation

Excel-DNA is responsible for loading the ribbon helper add-in, but is not otherwise involved in the ribbon extension. This means that the custom UI xml schema, and the signatures for the callback methods are exactly as documented by Microsoft. The best documentation for these aspects can be found in the three-part series on 'Customizing the 2007 Office Fluent Ribbon for Developers':

* [Part 1 - Overview](https://msdn.microsoft.com/en-us/library/aa338202.aspx)
* [Part 2 - Controls and callback reference](https://msdn.microsoft.com/en-us/library/aa338199.aspx)
* [Part 3 - Frequently asked questions, including C# and VB.NET callback signatures](https://msdn.microsoft.com/en-us/library/aa722523.aspx)

Information related to the Excel 2010 extensions to the ribbon can be found here:

* [Customizing Context Menus in Office 2010](https://msdn.microsoft.com/en-us/library/office/ee691832.aspx)
* [Customizing the Office 2010 Backstage View](https://msdn.microsoft.com/en-us/library/office/ee815851.aspx)
* [Ribbon Extensibility in Office 2010: Tab Activation and Auto-Scaling](https://msdn.microsoft.com/en-us/library/office/ee691834.aspx)

Other ribbon-related resources:

* [Ron de Bruin's Excel Tips](http://www.rondebruin.nl/win/s2/win003.htm)
* [Andy Pope's RibbonX Visual Designer](http://www.andypope.info/vba/ribboneditor.htm)

