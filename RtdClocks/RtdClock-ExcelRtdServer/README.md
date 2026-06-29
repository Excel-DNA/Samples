# RtdClock-ExcelRtdServer

This is the simplest `ExcelRtdServer` sample. It exposes `dnaRtdClock_ExcelRtdServer()`, which calls:

```csharp
XlCall.RTD(RtdClockServer.ServerProgId, null, "");
```

That uses Excel-DNA's on-demand RTD registration path, so the sample does not require explicit COM registration before the worksheet function is called.

For the advanced pre-registered COM server path, including direct worksheet formulas like:

```excel
=RTD("RtdClock.ClockServer",,"")
```

see the sibling sample `RtdClock-ExcelRtdServer-PreRegistered`.
