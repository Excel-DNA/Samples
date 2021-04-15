Imports ExcelDna.Integration
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel ' for event procedures...
Imports Microsoft.Vbe.Interop
Imports System.Runtime.InteropServices


''' <summary>AddIn Connection class, also handling Events from Excel (Open, Close, Activate)</summary>
<ComVisible(True)>
Public Class AddInEvents
    Implements IExcelAddIn

    ''' <summary>the app object needed for excel event handling (most of this class is dedicated to that)</summary>
    WithEvents Application As Excel.Application
    ''' <summary>CommandButton that can be inserted on a worksheet (name property being the same as the respective target range (for DBMapper/DBAction) or DBSeqnce Name)</summary>
    Shared WithEvents cb1 As Forms.CommandButton
    ''' <summary>CommandButton that can be inserted on a worksheet (name property being the same as the respective target range (for DBMapper/DBAction) or DBSeqnce Name)</summary>
    Shared WithEvents cb2 As Forms.CommandButton
    ''' <summary>CommandButton that can be inserted on a worksheet (name property being the same as the respective target range (for DBMapper/DBAction) or DBSeqnce Name)</summary>
    Shared WithEvents cb3 As Forms.CommandButton
    ''' <summary>CommandButton that can be inserted on a worksheet (name property being the same as the respective target range (for DBMapper/DBAction) or DBSeqnce Name)</summary>
    Shared WithEvents cb4 As Forms.CommandButton
    ''' <summary>CommandButton that can be inserted on a worksheet (name property being the same as the respective target range (for DBMapper/DBAction) or DBSeqnce Name)</summary>
    Shared WithEvents cb5 As Forms.CommandButton

    ''' <summary>connect to Excel when opening Addin</summary>
    Public Sub AutoOpen() Implements IExcelAddIn.AutoOpen
        Application = ExcelDnaUtil.Application
        ' for finding out what happened...
        Trace.Listeners.Add(New ExcelDna.Logging.LogDisplayTraceListener())
        ' Ribbon and context menu setup
        Globals.theMenuHandler = New MenuHandler
        Globals.LogInfo("initialize configuration settings")
        Globals.DBModifDefColl = New Dictionary(Of String, Dictionary(Of String, String)) From
            {{"DBMapper", New Dictionary(Of String, String) From {{"DBMapperTest1", "one"}, {"DBMapperTest2", "two"}}},
            {"DBAction", New Dictionary(Of String, String) From {{"DBActionTest1", "one"}, {"DBActionTest2", "two"}}},
            {"DBSeqnce", New Dictionary(Of String, String) From {{"DBSeqnceTest1", "one"}, {"DBSeqnceTest2", "two"}}}}
        ' get the ExcelDna LogDisplayTraceListener for filtering log messages by level in about box
        For Each srchdListener As Object In Trace.Listeners
            If srchdListener.ToString() = "ExcelDna.Logging.LogDisplayTraceListener" Then
                Globals.theLogListener = srchdListener
                Exit For
            End If
        Next
        ' initialize settings and get the default environment
        Globals.initSettings()
        ' Configs are 1 based, selectedEnvironment(index of environment dropdown) is 0 based. negative values not allowed!
        Dim selEnv As Integer = 1
    End Sub

    ''' <summary>AutoClose cleans up after finishing addin</summary>
    Public Sub AutoClose() Implements IExcelAddIn.AutoClose
        Try
            Globals.theMenuHandler = Nothing
        Catch ex As Exception
            Globals.ErrorMsg("Unloading error: " + ex.Message, "AutoClose")
        End Try
    End Sub

    ''' <summary>specific click handlers for the five definable commandbuttons</summary>
    Private Shared Sub cb1_Click() Handles cb1.Click
        cbClick(cb1.Name)
    End Sub
    Private Shared Sub cb2_Click() Handles cb2.Click
        cbClick(cb2.Name)
    End Sub
    Private Shared Sub cb3_Click() Handles cb3.Click
        cbClick(cb3.Name)
    End Sub
    Private Shared Sub cb4_Click() Handles cb4.Click
        cbClick(cb4.Name)
    End Sub
    Private Shared Sub cb5_Click() Handles cb5.Click
        cbClick(cb5.Name)
    End Sub

    ''' <summary>common click handler for all commandbuttons</summary>
    ''' <param name="cbName">name of command button, defines whether a DBModification is invoked (starts with DBMapper/DBAction/DBSeqnce)</param>
    Private Shared Sub cbClick(cbName As String)
        ' reset noninteractive messages (used for VBA invocations) and hadError for interactive invocations
        nonInteractiveErrMsgs = ""
        Dim DBModifType As String = Left(cbName, 8)
        If My.Computer.Keyboard.CtrlKeyDown And My.Computer.Keyboard.ShiftKeyDown Then
            MsgBox("editing DBModif " + cbName + " ...")
        Else
            MsgBox("DBmodif " + cbName + " activated...")
        End If
    End Sub

    ''' <summary>assign click handlers to commandbuttons in passed sheet Sh, maximum 5 buttons are supported</summary>
    ''' <param name="Sh"></param>
    Public Shared Function assignHandler(Sh As Object) As Boolean
        cb1 = Nothing : cb2 = Nothing : cb3 = Nothing : cb4 = Nothing : cb5 = Nothing
        assignHandler = True
        For Each shp As Excel.Shape In Sh.Shapes
            ' Associate clickhandler with all click events of the CommandButtons.
            Dim ctrlName As String
            Try : ctrlName = Sh.OLEObjects(shp.Name).Object.Name : Catch ex As Exception : ctrlName = "" : End Try
            If Left(ctrlName, 8) = "DBMapper" Or Left(ctrlName, 8) = "DBAction" Or Left(ctrlName, 8) = "DBSeqnce" Then
                If cb1 Is Nothing Then
                    cb1 = Sh.OLEObjects(shp.Name).Object
                ElseIf cb2 Is Nothing Then
                    cb2 = Sh.OLEObjects(shp.Name).Object
                ElseIf cb3 Is Nothing Then
                    cb3 = Sh.OLEObjects(shp.Name).Object
                ElseIf cb4 Is Nothing Then
                    cb4 = Sh.OLEObjects(shp.Name).Object
                ElseIf cb5 Is Nothing Then
                    cb5 = Sh.OLEObjects(shp.Name).Object
                Else
                    Globals.ErrorMsg("only max. of five DBModifier Buttons allowed on a Worksheet, currently using " + cb1.Name + "," + cb2.Name + "," + cb3.Name + "," + cb4.Name + " and " + cb5.Name + " !")
                    assignHandler = False
                    Exit For
                End If
            End If
        Next
    End Function

    Private WithEvents mInsertButton As Microsoft.Office.Core.CommandBarButton
    Private WithEvents mDeleteButton As Microsoft.Office.Core.CommandBarButton

    ' in case you need dynamic dependent context menus, you still have to resort to the old method of SheetBeforeRightClick and Application.CommandBars(...)
    Private Sub Application_SheetBeforeRightClick(Sh As Object, Target As Range, ByRef Cancel As Boolean) Handles Application.SheetBeforeRightClick
        ' check if we are in a DBMapper, if not then leave...
        If Globals.DBModifDefColl.ContainsKey("DBMapper") Then
            Dim targetName As String = getDBModifNameFromRange(Target)
            If Left(targetName, 8) <> "DBMapper" Then Exit Sub
        Else
            Exit Sub
        End If
        Dim appsCommandBars As String() = {"List Range Popup", "Row"}
        'first delete buttons
        For Each builtin As String In appsCommandBars
            Dim srchInsertButton = ExcelDnaUtil.Application.CommandBars(builtin).FindControl(Tag:="insTag")
            Dim srchDeleteButton = ExcelDnaUtil.Application.CommandBars(builtin).FindControl(Tag:="delTag")
            If Not srchInsertButton Is Nothing Then srchInsertButton.Delete()
            If Not srchDeleteButton Is Nothing Then srchDeleteButton.Delete()
        Next
        ' add context menus
        ' for whole sheet don't display DBSheet context menus !!
        If Not Target.Rows.Count = Target.EntireColumn.Rows.Count Then
            For Each builtin As String In appsCommandBars
                With ExcelDnaUtil.Application.CommandBars(builtin).Controls.Add(Type:=1, Before:=1, Temporary:=True)
                    .caption = "delete Row (Ctl-Sh-D)"
                    .FaceID = 214
                    .Tag = "delTag"
                End With
                With ExcelDnaUtil.Application.CommandBars(builtin).Controls.Add(Type:=1, Before:=1, Temporary:=True)
                    .caption = "insert Row (Ctl-Sh-I)"
                    .FaceID = 213
                    .Tag = "insTag"
                End With
            Next
        End If
        mInsertButton = ExcelDnaUtil.Application.CommandBars.FindControl(Tag:="insTag")
        mDeleteButton = ExcelDnaUtil.Application.CommandBars.FindControl(Tag:="delTag")
    End Sub

    ''' <summary>dynamic context menu item delete: delete row in CUD Style DBMappers</summary>
    ''' <param name="Ctrl"></param>
    ''' <param name="CancelDefault"></param>
    Private Sub mDeleteButton_Click(Ctrl As Microsoft.Office.Core.CommandBarButton, ByRef CancelDefault As Boolean) Handles mDeleteButton.Click
        MsgBox("Delete row!")
    End Sub

    ''' <summary>dynamic context menu item insert: insert row in CUD Style DBMappers</summary>
    ''' <param name="Ctrl"></param>
    ''' <param name="CancelDefault"></param>
    Private Sub mInsertButton_Click(Ctrl As Microsoft.Office.Core.CommandBarButton, ByRef CancelDefault As Boolean) Handles mInsertButton.Click
        MsgBox("Insert row!")
    End Sub
End Class