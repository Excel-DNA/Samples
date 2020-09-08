Imports ExcelDna.Integration

Namespace RtdClock_ExcelRtdServer
    Public Module RtdClock
        <ExcelFunction(Description:="Provides a ticking clock")>
        Public Function dnaRtdClock_ExcelRtdServer() As Object
            ' Call the Excel-DNA RTD wrapper, which does dynamic registration of the RTD server
            ' Note that the topic information needs at least one string - it's not used in this sample
            Return XlCall.RTD(RtdClockServer.ServerProgId, Nothing, "")
        End Function
    End Module
End Namespace
