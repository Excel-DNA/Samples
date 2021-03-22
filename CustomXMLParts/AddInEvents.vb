Imports ExcelDna.Integration
Imports Microsoft.Office.Interop
Imports System.Runtime.InteropServices

''' <summary>AddIn Connection class, also handling Events from Excel (Open, Close, Activate)</summary>
<ComVisible(True)>
Public Class AddInEvents
    Implements IExcelAddIn

    WithEvents Application As Excel.Application

    ''' <summary>connect to Excel when opening Addin</summary>
    Public Sub AutoOpen() Implements IExcelAddIn.AutoOpen
        Application = ExcelDnaUtil.Application

        ' Ribbon menu setup
        DBModifs.theMenuHandler = New MenuHandler
        DBModifs.DBModifDefColl = New Dictionary(Of String, Dictionary(Of String, DBModif))
    End Sub

    ''' <summary>AutoClose cleans up after finishing addin</summary>
    Public Sub AutoClose() Implements IExcelAddIn.AutoClose
        Try
            DBModifs.theMenuHandler = Nothing
        Catch ex As Exception
            DBModifs.ErrorMsg("DBAddin unloading error: " + ex.Message, "AutoClose")
        End Try
    End Sub

    ''' <summary>WorkbookActivate: gets defined named ranges for DBMapper invocation in the current workbook after activation and updates Ribbon with it</summary>
    Private Sub Application_WorkbookActivate(Wb As Excel.Workbook) Handles Application.WorkbookActivate
        If DBModifs.DBModifDefColl Is Nothing Then
            DBModifs.DBModifDefColl = New Dictionary(Of String, Dictionary(Of String, DBModif))
        End If
        DBModifs.getDBModifDefinitions()
    End Sub

    Private WBclosing As Boolean
    Private Sub Application_WorkbookBeforeClose(Wb As Excel.Workbook, ByRef Cancel As Boolean) Handles Application.WorkbookBeforeClose
        WBclosing = True
    End Sub

    ' don't remove DBModifDefColl in Application_WorkbookBeforeClose, as we might cancel the close. Only when the Workbook becomes deactivated, it's really closed.
    Private Sub Application_WorkbookDeactivate(Wb As Excel.Workbook) Handles Application.WorkbookDeactivate
        If DBModifs.DBModifDefColl.Count > 0 And WBclosing Then
            DBModifs.DBModifDefColl.Clear()
            DBModifs.theRibbon.Invalidate()
        End If
        WBclosing = False
    End Sub
End Class