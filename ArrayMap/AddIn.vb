Imports ExcelDna.Integration
Imports ExcelDna.IntelliSense

Public Class AddIn
    Implements IExcelAddIn

    Public Sub AutoOpen() Implements IExcelAddIn.AutoOpen
        IntelliSenseServer.Install()
    End Sub

    Public Sub AutoClose() Implements IExcelAddIn.AutoClose
        IntelliSenseServer.Uninstall()
    End Sub
End Class
