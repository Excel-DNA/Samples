Imports ExcelDna.Integration
Imports Microsoft.Office.Interop

''' <summary>Global variables and functions for DB Addin</summary>
Public Module Globals
    ' general Global objects/variables
    ''' <summary>ribbon menu handler</summary>
    Public theMenuHandler As MenuHandler
    ''' <summary>currently selected environment for DB Functions, zero based (env -1) !!</summary>
    Public selectedEnvironment As Integer
    ''' <summary>reference object for the Addins ribbon</summary>
    Public theRibbon As CustomUI.IRibbonUI
    ''' <summary>environment definitions</summary>
    Public environdefs As String() = {}
    ''' <summary>DBModif definition collections of DBmodif types (key of top level dictionary) with values beinig collections of DBModifierNames (key of contained dictionaries) and DBModifiers (value of contained dictionaries))</summary>
    Public DBModifDefColl As Dictionary(Of String, Dictionary(Of String, String))
    ''' <summary>the selected event level in the About box</summary>
    Public EventLevelSelected As String
    ''' <summary>the log listener</summary>
    Public theLogListener As TraceListener
    ''' <summary>for DBMapper invocations by execDBModif, this is set to true, avoiding MsgBox</summary>
    Public nonInteractive As Boolean = False
    ''' <summary>collect non interactive error messages here</summary>
    Public nonInteractiveErrMsgs As String
    ''' <summary>set to true if warning was issued</summary>
    Public WarningIssued As Boolean
    Public DontChangeEnvironment As Boolean

    ' Global settings
    Public DebugAddin As Boolean
    ''' <summary>Default ConnectionString, if no connection string is given by user....</summary>
    Public ConstConnString As String
    ''' <summary>global connection timeout (can't be set in DB functions)</summary>
    Public CnnTimeout As Integer
    ''' <summary>global command timeout (can't be set in DB functions)</summary>
    Public CmdTimeout As Integer
    ''' <summary>default formatting style used in DBDate</summary>
    Public DefaultDBDateFormatting As Integer

    ''' <summary>environment for settings (+1 of selectedeEnvironment which is the index of the dropdown)</summary>
    ''' <returns></returns>
    Public Function env() As String
        Return (Globals.selectedEnvironment + 1).ToString()
    End Function

    ''' <summary>initializes global configuration variables</summary>
    Public Sub initSettings()
        Try
            ' load environments
            Dim i As Integer = 1
            ReDim Preserve environdefs(-1)
            Do
                ReDim Preserve environdefs(environdefs.Length)
                environdefs(environdefs.Length - 1) = "Environment - " + i.ToString()
                i += 1
            Loop Until i = 7
        Catch ex As Exception
            ErrorMsg("Error in initialization of Settings: " + ex.Message)
        End Try
    End Sub

    ''' <summary>Logs Message of eEventType to System.Diagnostics.Trace</summary>
    ''' <param name="Message">Message to be logged</param>
    ''' <param name="eEventType">event type: info, warning, error</param>
    ''' <param name="caller">reflection based caller information: module.method</param>
    Private Sub WriteToLog(Message As String, eEventType As EventLogEntryType, caller As String)
        ' collect errors and warnings for returning messages in executeDBModif
        If eEventType = EventLogEntryType.Error Or eEventType = EventLogEntryType.Warning Then nonInteractiveErrMsgs += caller + ":" + Message + vbCrLf
        If nonInteractive Then
            Trace.TraceInformation("Noninteractive: {0}: {1}", caller, Message)
        Else
            Select Case eEventType
                Case EventLogEntryType.Information
                    Trace.TraceInformation("{0}: {1}", caller, Message)
                Case EventLogEntryType.Warning
                    Trace.TraceWarning("{0}: {1}", caller, Message)
                    WarningIssued = True
                    ' at Addin Start ribbon has not been loaded so avoid call to it here..
                    If Not theRibbon Is Nothing Then theRibbon.InvalidateControl("showLog")
                Case EventLogEntryType.Error
                    Trace.TraceError("{0}: {1}", caller, Message)
            End Select
        End If
    End Sub

    ''' <summary>Logs error messages</summary>
    ''' <param name="LogMessage">the message to be logged</param>
    Public Sub LogError(LogMessage As String)
        Dim theMethod As Object = (New System.Diagnostics.StackTrace).GetFrame(1).GetMethod
        Dim caller As String = theMethod.ReflectedType.FullName + "." + theMethod.Name
        WriteToLog(LogMessage, EventLogEntryType.Error, caller)
    End Sub

    ''' <summary>Logs warning messages</summary>
    ''' <param name="LogMessage">the message to be logged</param>
    Public Sub LogWarn(LogMessage As String)
        Dim theMethod As Object = (New System.Diagnostics.StackTrace).GetFrame(1).GetMethod
        Dim caller As String = theMethod.ReflectedType.FullName + "." + theMethod.Name
        WriteToLog(LogMessage, EventLogEntryType.Warning, caller)
    End Sub

    ''' <summary>Logs informational messages</summary>
    ''' <param name="LogMessage">the message to be logged</param>
    Public Sub LogInfo(LogMessage As String)
        If DebugAddin Then
            Dim theMethod As Object = (New System.Diagnostics.StackTrace).GetFrame(1).GetMethod
            Dim caller As String = theMethod.ReflectedType.FullName + "." + theMethod.Name
            WriteToLog(LogMessage, EventLogEntryType.Information, caller)
        End If
    End Sub

    ''' <summary>show Error message to User and log as warning (errors would pop up the trace information window)</summary> 
    ''' <param name="LogMessage">the message to be shown/logged</param>
    ''' <param name="errTitle">optionally pass a title for the msgbox here</param>
    Public Sub ErrorMsg(LogMessage As String, Optional errTitle As String = "RibbonVB Error", Optional msgboxIcon As MsgBoxStyle = MsgBoxStyle.Critical)
        Dim theMethod As Object = (New System.Diagnostics.StackTrace).GetFrame(1).GetMethod
        Dim caller As String = theMethod.ReflectedType.FullName + "." + theMethod.Name
        WriteToLog(LogMessage, If(msgboxIcon = MsgBoxStyle.Critical Or msgboxIcon = MsgBoxStyle.Exclamation, EventLogEntryType.Warning, EventLogEntryType.Information), caller) ' to avoid popup of trace log in nonInteractive mode...
        If Not nonInteractive Then MsgBox(LogMessage, msgboxIcon + MsgBoxStyle.OkOnly, errTitle)
    End Sub

    Public Function QuestionMsg(theMessage As String, Optional questionType As MsgBoxStyle = MsgBoxStyle.OkCancel, Optional questionTitle As String = "RibbonVB Question", Optional msgboxIcon As MsgBoxStyle = MsgBoxStyle.Question) As MsgBoxResult
        Dim theMethod As Object = (New System.Diagnostics.StackTrace).GetFrame(1).GetMethod
        Dim caller As String = theMethod.ReflectedType.FullName + "." + theMethod.Name
        WriteToLog(theMessage, If(msgboxIcon = MsgBoxStyle.Critical Or msgboxIcon = MsgBoxStyle.Exclamation, EventLogEntryType.Warning, EventLogEntryType.Information), caller) ' to avoid popup of trace log
        If nonInteractive Then
            If questionType = MsgBoxStyle.OkCancel Then Return MsgBoxResult.Cancel
            If questionType = MsgBoxStyle.YesNo Then Return MsgBoxResult.No
            If questionType = MsgBoxStyle.YesNoCancel Then Return MsgBoxResult.No
            If questionType = MsgBoxStyle.RetryCancel Then Return MsgBoxResult.Cancel
        End If
        Return MsgBox(theMessage, msgboxIcon + questionType, questionTitle)
    End Function

    ''' <summary>gets DB Modification Name (DBMapper or DBAction) from theRange</summary>
    ''' <param name="theRange"></param>
    ''' <returns>the retrieved name as a string (not name object !)</returns>
    Public Function getDBModifNameFromRange(theRange As Excel.Range) As String
        Dim nm As Excel.Name
        Dim rng, testRng As Excel.Range

        getDBModifNameFromRange = ""
        If theRange Is Nothing Then Exit Function
        Try
            ' try all names in workbook
            For Each nm In theRange.Parent.Parent.Names
                rng = Nothing
                ' test whether range referring to that name (if it is a real range)...
                Try : rng = nm.RefersToRange : Catch ex As Exception : End Try
                If Not rng Is Nothing Then
                    testRng = Nothing
                    ' ...intersects with the passed range
                    Try : testRng = ExcelDnaUtil.Application.Intersect(theRange, rng) : Catch ex As Exception : End Try
                    If Not testRng Is Nothing And (InStr(1, nm.Name, "DBMapper") >= 1 Or InStr(1, nm.Name, "DBAction") >= 1) Then
                        ' and pass back the name if it does and is a DBMapper or a DBAction
                        getDBModifNameFromRange = nm.Name
                        Exit Function
                    End If
                End If
            Next
        Catch ex As Exception
            ErrorMsg("Exception: " + ex.Message, "get DBModif Name From Range")
        End Try
    End Function

    ' shortcut for context menu button 1
    <ExcelCommand(Name:="button1", ShortCut:="^R")>
    Public Sub button1()
        MsgBox("context menu button 1/shortcut Ctrl-Shift R clicked...")
    End Sub

    ' shortcut for context menu button 2
    <ExcelCommand(Name:="button2", ShortCut:="^J")>
    Public Sub button2()
        MsgBox("context menu button 2/shortcut Ctrl-Shift J clicked...")
    End Sub

End Module