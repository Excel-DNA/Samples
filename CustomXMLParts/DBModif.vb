Imports ExcelDna.Integration
Imports Microsoft.Office.Core

''' <summary>DBModifs are used to store a range of data in Excel to the database, do a DB Action or a sequence of DBMappers and DBActions</summary>
Public Class DBModif
    '''<summary>unique key of DBModif</summary>
    Private dbmodifName As String
    '''<summary>should DBMap be saved / DBAction be done on Excel Saving? (default no)</summary>
    Private execOnSave As Boolean = False
    ''' <summary>ask for confirmation before executtion of DBModif</summary>
    Private askBeforeExecute As Boolean = True
    ''' <summary>environment specific for the DBModif object, if left empty then set to default environment (either 0 or currently selected environment)</summary>
    Private env As String = ""
    ''' <summary>Text displayed for confirmation before doing dbModif instead of standard text</summary>
    Private confirmText As String = ""
    ''' <summary>Database to store to</summary>
    Private database As String
    ''' <summary>Database Table, where Data is to be stored</summary>
    Private tableName As String = ""
    ''' <summary>count of primary keys in datatable, starting from the leftmost column</summary>
    Private primKeysCount As Integer = 0
    ''' <summary>if set, then insert row into table if primary key is missing there. Default = False (only update)</summary>
    Private insertIfMissing As Boolean = False
    ''' <summary>additional stored procedure to be executed after saving</summary>
    Private executeAdditionalProc As String = ""
    ''' <summary>columns to be ignored (helper columns)</summary>
    Private ignoreColumns As String = ""
    ''' <summary>respect C/U/D Flags (DBSheet functionality)</summary>
    Public CUDFlags As Boolean = False
    ''' <summary>if set, don't notify error values in cells during update/insert</summary>
    Private IgnoreDataErrors As Boolean = False
    '''<summary>first columnn is treated as an autoincrementing key column</summary>
    Private AutoIncFlag As Boolean = False
    ''' <summary>sequence of DB Mappers, DB Actions and DB Refreshes being executed in this sequence</summary>
    Private sequenceParams() As String = {}

    Public Sub New(definitionXML As CustomXMLNode, DBModifType As String)
        ' if we have a Name attribute use this, if not, use the basename of the type + "unknown"
        If definitionXML.Attributes.Count > 0 Then
            dbmodifName = definitionXML.Attributes(1).Text
        Else
            dbmodifName = definitionXML.BaseName + "unknown"
        End If
        execOnSave = Convert.ToBoolean(getParamFromXML(definitionXML, "execOnSave", "Boolean"))
        askBeforeExecute = Convert.ToBoolean(getParamFromXML(definitionXML, "askBeforeExecute", "Boolean"))
        confirmText = getParamFromXML(definitionXML, "confirmText")

        If DBModifType = "DBMapper" Then
            ' fill parameters from definition
            env = getParamFromXML(definitionXML, "env")
            database = getParamFromXML(definitionXML, "database")
            If database = "" Then Throw New Exception("No database given in DBMapper definition!")
            tableName = getParamFromXML(definitionXML, "tableName")
            If tableName = "" Then Throw New Exception("No Tablename given in DBMapper definition!")
            Try
                primKeysCount = Convert.ToInt32(getParamFromXML(definitionXML, "primKeysStr"))
            Catch ex As Exception
                Throw New Exception("couldn't get primary key count given in DBMapper definition:" + ex.Message)
            End Try
            insertIfMissing = Convert.ToBoolean(getParamFromXML(definitionXML, "insertIfMissing", "Boolean"))
            executeAdditionalProc = getParamFromXML(definitionXML, "executeAdditionalProc")
            ignoreColumns = getParamFromXML(definitionXML, "ignoreColumns")
            IgnoreDataErrors = Convert.ToBoolean(getParamFromXML(definitionXML, "IgnoreDataErrors", "Boolean"))
            CUDFlags = Convert.ToBoolean(getParamFromXML(definitionXML, "CUDFlags", "Boolean"))
            AutoIncFlag = Convert.ToBoolean(getParamFromXML(definitionXML, "AutoIncFlag", "Boolean"))
        ElseIf DBModifType = "DBAction" Then
            env = getParamFromXML(definitionXML, "env")
            database = getParamFromXML(definitionXML, "database")
            If database = "" Then Throw New Exception("No database given in DBAction definition!")
        ElseIf DBModifType = "DBSeqnce" Then
            Dim seqSteps As Integer = definitionXML.SelectNodes("ns0:seqStep").Count
            If seqSteps = 0 Then
                Throw New Exception("no steps defined in DBSequence definition!")
            Else
                ReDim sequenceParams(seqSteps - 1)
                For i = 1 To seqSteps
                    sequenceParams(i - 1) = definitionXML.SelectNodes("ns0:seqStep")(i).Text
                Next
            End If
        End If
    End Sub

    Public Sub displayDBModif()
        Dim displayedData As String = ""
        displayedData += "dbmodifName:" + dbmodifName + vbCrLf
        displayedData += "execOnSave:" + execOnSave.ToString() + vbCrLf
        displayedData += "askBeforeExecute:" + askBeforeExecute.ToString() + vbCrLf
        displayedData += "confirmText:" + confirmText + vbCrLf

        If Left(dbmodifName, 8) = "DBMapper" Then
            ' fill parameters from definition
            displayedData += "env:" + env + vbCrLf
            displayedData += "database:" + database + vbCrLf
            displayedData += "tableName:" + tableName + vbCrLf
            displayedData += "primKeysCount:" + primKeysCount.ToString() + vbCrLf
            displayedData += "insertIfMissing:" + insertIfMissing.ToString() + vbCrLf
            displayedData += "executeAdditionalProc:" + executeAdditionalProc + vbCrLf
            displayedData += "ignoreColumns:" + ignoreColumns + vbCrLf
            displayedData += "IgnoreDataErrors:" + IgnoreDataErrors.ToString() + vbCrLf
            displayedData += "CUDFlags:" + CUDFlags.ToString() + vbCrLf
            displayedData += "AutoIncFlag:" + AutoIncFlag.ToString() + vbCrLf
        ElseIf Left(dbmodifName, 8) = "DBAction" Then
            displayedData += "env:" + env + vbCrLf
            displayedData += "database:" + database + vbCrLf
        ElseIf Left(dbmodifName, 8) = "DBSeqnce" Then
            For i = 1 To sequenceParams.Length
                displayedData += "sequenceStep" + i.ToString() + ":" + sequenceParams(i - 1) + vbCrLf
            Next
        End If
        MsgBox(displayedData, MsgBoxStyle.Information, "CustomXMLParts")
    End Sub

    ''' <summary>wrapper to get the single definition element values from the DBModifier CustomXML node, also checks for multiple definition elements</summary>
    ''' <param name="definitionXML">the CustomXML node for the DBModifier</param>
    ''' <param name="nodeName">the definition element's name (eg "env")</param>
    ''' <returns>the definition element's value</returns>
    ''' <exception cref="Exception">if multiple elements exist for the definition element's name throw warning !</exception>
    Protected Function getParamFromXML(definitionXML As CustomXMLNode, nodeName As String, Optional ReturnType As String = "") As String
        Dim nodeCount As Integer = definitionXML.SelectNodes("ns0:" + nodeName).Count
        If nodeCount = 0 Then
            getParamFromXML = "" ' optional nodes become empty strings
        Else
            getParamFromXML = definitionXML.SelectSingleNode("ns0:" + nodeName).Text
        End If
        If ReturnType = "Boolean" And getParamFromXML = "" Then getParamFromXML = "False"
    End Function

End Class


''' <summary>global helper functions for DBModifiers</summary>
Public Module DBModifs
    ' general Global objects/variables
    ''' <summary>ribbon menu handler</summary>
    Public theMenuHandler As MenuHandler
    ''' <summary>reference object for the Addins ribbon</summary>
    Public theRibbon As CustomUI.IRibbonUI
    ''' <summary>DBModif definition collections of DBmodif types (key of top level dictionary) with values beinig collections of DBModifierNames (key of contained dictionaries) and DBModifiers (value of contained dictionaries))</summary>
    Public DBModifDefColl As Dictionary(Of String, Dictionary(Of String, DBModif))

    ''' <summary>show Error message to User and log as warning (errors would pop up the trace information window)</summary> 
    ''' <param name="LogMessage">the message to be shown/logged</param>
    ''' <param name="errTitle">optionally pass a title for the msgbox here</param>
    Public Sub ErrorMsg(LogMessage As String, Optional errTitle As String = "CustomXMLParts Error", Optional msgboxIcon As MsgBoxStyle = MsgBoxStyle.Critical)
        MsgBox(LogMessage, msgboxIcon + MsgBoxStyle.OkOnly, errTitle)
    End Sub

    Public Function QuestionMsg(theMessage As String, Optional questionType As MsgBoxStyle = MsgBoxStyle.OkCancel, Optional questionTitle As String = "CustomXMLParts Question", Optional msgboxIcon As MsgBoxStyle = MsgBoxStyle.Question) As MsgBoxResult
        Return MsgBox(theMessage, msgboxIcon + questionType, questionTitle)
    End Function

    ''' <summary>creates a DBModifier in the CustomXMLNode</summary>
    Public Sub createDBModif(createdDBModifType As String)

        Dim CustomXmlParts As Object = ExcelDnaUtil.Application.ActiveWorkbook.CustomXMLParts.SelectByNamespace("DBModifDef")
        If CustomXmlParts.Count = 0 Then
            ' in case no CustomXmlPart in Namespace DBModifDef exists in the workbook, add one
            ExcelDnaUtil.Application.ActiveWorkbook.CustomXMLParts.Add("<root xmlns=""DBModifDef""></root>")
            CustomXmlParts = ExcelDnaUtil.Application.ActiveWorkbook.CustomXMLParts.SelectByNamespace("DBModifDef")
        End If

        ' NamespaceURI:="DBModifDef" is required to avoid adding a xmlns attribute to each element.
        CustomXmlParts(1).SelectSingleNode("/ns0:root").AppendChildNode(createdDBModifType, NamespaceURI:="DBModifDef")
        ' new appended elements are last, get it to append further child elements
        Dim dbModifNode As CustomXMLNode = CustomXmlParts(1).SelectSingleNode("/ns0:root").LastChild
        ' append the detailed settings to the definition element
        dbModifNode.AppendChildNode("Name", NodeType:=MsoCustomXMLNodeType.msoCustomXMLNodeAttribute, NodeValue:=createdDBModifType + Guid.NewGuid().ToString())
        dbModifNode.AppendChildNode("execOnSave", NamespaceURI:="DBModifDef", NodeValue:="True")
        dbModifNode.AppendChildNode("askBeforeExecute", NamespaceURI:="DBModifDef", NodeValue:="True")
        If createdDBModifType = "DBMapper" Then
            dbModifNode.AppendChildNode("env", NamespaceURI:="DBModifDef", NodeValue:="1") ' if not selected, set environment to 0 (default anyway)
            dbModifNode.AppendChildNode("database", NamespaceURI:="DBModifDef", NodeValue:="SomeDB")
            dbModifNode.AppendChildNode("tableName", NamespaceURI:="DBModifDef", NodeValue:="SomeTable")
            dbModifNode.AppendChildNode("primKeysStr", NamespaceURI:="DBModifDef", NodeValue:="2")
            dbModifNode.AppendChildNode("insertIfMissing", NamespaceURI:="DBModifDef", NodeValue:="True")
            dbModifNode.AppendChildNode("executeAdditionalProc", NamespaceURI:="DBModifDef", NodeValue:="SomeStoredProc")
            dbModifNode.AppendChildNode("ignoreColumns", NamespaceURI:="DBModifDef", NodeValue:="ignoredCol1,ignoredCol2")
            dbModifNode.AppendChildNode("CUDFlags", NamespaceURI:="DBModifDef", NodeValue:="True")
            dbModifNode.AppendChildNode("AutoIncFlag", NamespaceURI:="DBModifDef", NodeValue:="False")
            dbModifNode.AppendChildNode("IgnoreDataErrors", NamespaceURI:="DBModifDef", NodeValue:="False")
        ElseIf createdDBModifType = "DBAction" Then
            dbModifNode.AppendChildNode("env", NamespaceURI:="DBModifDef", NodeValue:="1")
            dbModifNode.AppendChildNode("database", NamespaceURI:="DBModifDef", NodeValue:="SomeDB")
        ElseIf createdDBModifType = "DBSeqnce" Then
            dbModifNode.AppendChildNode("seqStep", NamespaceURI:="DBModifDef", NodeValue:="DBBegin:Begins DB Transaction")
            dbModifNode.AppendChildNode("seqStep", NamespaceURI:="DBModifDef", NodeValue:="DBMapper:Some Unknown Mapper")
            dbModifNode.AppendChildNode("seqStep", NamespaceURI:="DBModifDef", NodeValue:="DBCommitRollback:Commits or Rolls back DB Transaction")
        End If
        dbModifNode.AppendChildNode("confirmText", NamespaceURI:="DBModifDef", NodeValue:="additionally displayed confirm text")
    End Sub

    ' load DBModifier definitions as objects into Global collection DBModifDefColl
    Public Sub getDBModifDefinitions()
        Try
            DBModifs.DBModifDefColl.Clear()
            Dim CustomXmlParts As Object = ExcelDnaUtil.Application.ActiveWorkbook.CustomXMLParts.SelectByNamespace("DBModifDef")
            If CustomXmlParts.Count = 1 Then
                ' read definitions from CustomXMLParts
                For Each customXMLNodeDef As CustomXMLNode In CustomXmlParts(1).SelectSingleNode("/ns0:root").ChildNodes
                    Dim DBModiftype As String = Left(customXMLNodeDef.BaseName, 8)
                    If DBModiftype = "DBSeqnce" Or DBModiftype = "DBMapper" Or DBModiftype = "DBAction" Then
                        Dim nodeName As String
                        If customXMLNodeDef.Attributes.Count > 0 Then
                            nodeName = customXMLNodeDef.Attributes(1).Text
                        Else
                            nodeName = customXMLNodeDef.BaseName + "unknown"
                        End If

                        ' finally create the DBModif Object and fill parameters into CustomXMLPart:
                        Dim newDBModif As DBModif = New DBModif(customXMLNodeDef, DBModiftype)
                        ' ... and add it to the collection DBModifDefColl
                        Dim defColl As Dictionary(Of String, DBModif) ' definition lookup collection for DBModifiername -> object
                        If Not newDBModif Is Nothing Then
                            If Not DBModifDefColl.ContainsKey(DBModiftype) Then
                                ' add to new DBModiftype "menu"
                                defColl = New Dictionary(Of String, DBModif) From {
                                    {nodeName, newDBModif}
                                }
                                DBModifDefColl.Add(DBModiftype, defColl)
                            Else
                                ' add definition to existing DBModiftype "menu"
                                defColl = DBModifDefColl(DBModiftype)
                                If defColl.ContainsKey(nodeName) Then
                                    ErrorMsg("DBModifier " + nodeName + " already existing!", "get DBModif Definitions")
                                Else
                                    defColl.Add(nodeName, newDBModif)
                                End If
                            End If
                        End If
                    End If
EndOuterLoop:
                Next
            ElseIf CustomXmlParts.Count > 1 Then
                ErrorMsg("Multiple CustomXmlParts for DBModifDef existing!", "get DBModif Definitions")
            End If
            DBModifs.theRibbon.Invalidate()
        Catch ex As Exception
            ErrorMsg("Exception:  " + ex.Message, "get DBModif Definitions")
        End Try
    End Sub

End Module