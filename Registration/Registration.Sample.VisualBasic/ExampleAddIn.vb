﻿Imports ExcelDna.Integration

Public Class ExampleAddIn
    Implements IExcelAddIn

    Public Sub AutoOpen() Implements IExcelAddIn.AutoOpen
        ExcelIntegration.RegisterUnhandledExceptionHandler(Function(ex) "!!! ERROR: " + ex.ToString())
    End Sub

    Public Sub AutoClose() Implements IExcelAddIn.AutoClose

    End Sub

End Class
