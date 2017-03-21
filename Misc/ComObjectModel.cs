// Shows how to get hold of the Excel COM Application object

// Install the 'ExcelDna.Interop' package from NuGet, or reference the two assemblies:
// * Microsoft.Office.Interop.Excel.dll
// * Office.dll

using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;

public class TestCommands
{
    // Defines a macro that uses the COM object model to add a diamond shape to the active sheet.
    [ExcelCommand(ShortCut = "^D")] // Ctrl+Shift+D
    public static void AddDiamond()
    {
        Application xlApp = (Application)ExcelDnaUtil.Application;
        string version = xlApp.Version;

        Worksheet ws = xlApp.ActiveSheet as Worksheet;  // Need to change type - it might be a Chart, then ws would be null
        Shape diamond = ws.Shapes.AddShape(Microsoft.Office.Core.MsoAutoShapeType.msoShapeDiamond, 10, 10, 100, 100);
        diamond.Fill.BackColor.RGB = (int)XlRgbColor.rgbSlateBlue;
    }
}
