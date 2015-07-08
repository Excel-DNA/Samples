# Simple RTD server with array function wrapper

This sample was created to provide a simple RTD server to check the behaviour of RTD functions when called from array formulas.

The wrapper function has signature
   ```cs
	public static object RtdArrayTest(string prefix)
   ```
and returns a 2x1 array.

The function can be called from Excel as an array formula (with Ctrl+Shift+Enter) using another cell as the "prefix":
   ```
    {=RtdArrayTest(D1)}
   ```

The implementation of the RTD server is based on the Excel-DNA base class `ExcelRtdServer`, and just uses a Timer to update the topics.
