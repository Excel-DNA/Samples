# Ribbon Sample

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
