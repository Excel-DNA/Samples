Imports ExcelDna.Integration
Imports ExcelDna.ComInterop
Imports System.Runtime.InteropServices

<ClassInterface(ClassInterfaceType.AutoDispatch)>
<ProgId("ComAddin.FunctionLibrary")>
<ComVisible(True)>
Public Class AccessibleFunctions

    Public Function add(x As Double, y As Double)
        Return x + y
    End Function

End Class

<ComVisible(False)>
Public Class AddInEvents
    Implements IExcelAddIn

    Public Sub AutoOpen() Implements IExcelAddIn.AutoOpen
        ComServer.DllRegisterServer()
    End Sub

    Public Sub AutoClose() Implements IExcelAddIn.AutoClose
        ComServer.DllUnregisterServer()
    End Sub
End Class