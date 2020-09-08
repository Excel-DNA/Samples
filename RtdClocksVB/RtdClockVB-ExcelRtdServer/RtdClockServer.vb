Imports System.Collections.Generic
Imports System.Runtime.InteropServices
Imports System.Threading
Imports ExcelDna.Integration.Rtd

Namespace RtdClock_ExcelRtdServer
    <ComVisible(True)>                   ' Required since the default template puts [assembly:ComVisible(false)] in the AssemblyInfo.cs
    <ProgId(RtdClockServer.ServerProgId)>     '  If ProgId is not specified, change the XlCall.RTD call in the wrapper to use namespace + type name (the default ProgId)
    Public Class RtdClockServer
        Inherits ExcelRtdServer

        Public Const ServerProgId As String = "RtdClockVB.ClockServer"

        ' Using a System.Threading.Time which invokes the callback on a ThreadPool thread 
        ' (normally that would be dangeours for an RTD server, but ExcelRtdServer is thrad-safe)
        Private _timer As Timer
        Private _topics As List(Of Topic)

        Protected Overrides Function ServerStart() As Boolean
            _timer = New Timer(AddressOf timer_tick, Nothing, 0, 1000)
            _topics = New List(Of Topic)()
            Return True
        End Function

        Protected Overrides Sub ServerTerminate()
            _timer.Dispose()
        End Sub

        Protected Overrides Function ConnectData(ByVal topic As Topic, ByVal topicInfo As IList(Of String), ByRef newValues As Boolean) As Object
            _topics.Add(topic)
            Return Date.Now.ToString("HH:mm:ss") & " (ConnectData)"
        End Function

        Protected Overrides Sub DisconnectData(ByVal topic As Topic)
            _topics.Remove(topic)
        End Sub

        Private Sub timer_tick(ByVal _unused_state_ As Object)
            Dim now As String = Date.Now.ToString("HH:mm:ss")

            For Each topic In _topics
                topic.UpdateValue(now)
            Next
        End Sub
    End Class
End Namespace