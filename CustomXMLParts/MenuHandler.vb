Imports ExcelDna.Integration
Imports System.Runtime.InteropServices

''' <summary>handles all Menu related aspects</summary>
<ComVisible(True)>
Public Class MenuHandler
    Inherits CustomUI.ExcelRibbon

    ''' <summary>callback after Excel loaded the Ribbon, used to initialize data for the Ribbon</summary>
    Public Sub ribbonLoaded(theRibbon As CustomUI.IRibbonUI)
        DBModifs.theRibbon = theRibbon
    End Sub

    ''' <summary>creates the Ribbon (only at startup). Further changes to the ribbon can be done via dynamic menus and ribbon.invalidate</summary>
    ''' <returns></returns>
    Public Overrides Function GetCustomUI(RibbonID As String) As String
        ' Ribbon definition XML
        Dim customUIXml As String = "<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui' onLoad='ribbonLoaded' ><ribbon><tabs><tab id='AddinTab' label='CustomXMLParts'>"
        ' DBModif Group: maximum three DBModif types possible (depending on existence in current workbook): 
        customUIXml +=
        "<group id='DBModifGroup' label='Execute DBModifier'>"
        For Each DBModifType As String In {"DBSeqnce", "DBMapper", "DBAction"}
            customUIXml += "<dynamicMenu id='" + DBModifType + "' " +
                                                "size='large' getLabel='getDBModifTypeLabel' imageMso='ApplicationOptionsDialog' " +
                                                "getScreentip='getDBModifScreentip' getContent='getDBModifMenuContent' getVisible='getDBModifMenuVisible'/>"
        Next
        customUIXml += "<dialogBoxLauncher><button id='DBModifEdit' label='DBModif design' onAction='showDBModifEdit' screentip='Show/edit DBModif Definitions of current workbook'/></dialogBoxLauncher>"
        customUIXml += "</group>"
        customUIXml +=
            "<group id='DBModifAdd' label='Add DBModifier'>" +
            "<button id='add' label='Add DBModif' size='large' onAction='addDBModif' imageMso='ControlProperties' screentip='add DB Modifier' />" +
            "</group>"
        customUIXml += "</tab></tabs></ribbon></customUI>"
        Return customUIXml
    End Function

#Disable Warning IDE0060 ' Hide not used Parameter warning as this is very often the case with the below callbacks from the ribbon

    ''' <summary>show DBModif definitions edit box</summary>
    ''' <param name="control"></param>
    Sub showDBModifEdit(control As CustomUI.IRibbonControl)
        ' only show dialog if there is a workbook and it has the relevant custom XML part.
        If ExcelDnaUtil.Application.ActiveWorkbook Is Nothing Then
            DBModifs.ErrorMsg("No active workbook to store CustomXMLParts into. Create one, please!")
            Exit Sub
        ElseIf ExcelDnaUtil.Application.ActiveWorkbook.CustomXMLParts.SelectByNamespace("DBModifDef").Count > 0 Then
            Dim theEditDBModifDefDlg As EditDBModifDef = New EditDBModifDef()
            If theEditDBModifDefDlg.ShowDialog() = System.Windows.Forms.DialogResult.OK Then DBModifs.getDBModifDefinitions()
        Else
            If DBModifs.QuestionMsg("should a DBModif definition (DBMapper) be created?") = MsgBoxResult.Ok Then
                DBModifs.createDBModif("DBMapper")
            End If
        End If
        DBModifs.getDBModifDefinitions()
        theRibbon.Invalidate()
    End Sub

    ''' <summary>set the name of the DBModifType dropdown to the sheet name (for the WB dropdown this is the WB name)</summary>
    ''' <param name="control"></param>
    ''' <returns></returns>
    Public Function getDBModifTypeLabel(control As CustomUI.IRibbonControl) As String
        getDBModifTypeLabel = If(control.Id = "DBSeqnce", "DBSequence", control.Id)
    End Function

    ''' <summary>create the buttons in the DBModif sheet dropdown menu</summary>
    ''' <param name="control"></param>
    ''' <returns>the menu content xml</returns>
    Public Function getDBModifMenuContent(control As CustomUI.IRibbonControl) As String
        Dim xmlString As String = "<menu xmlns='http://schemas.microsoft.com/office/2009/07/customui'>"
        Try
            If Not DBModifs.DBModifDefColl.ContainsKey(control.Id) Then Return ""
            For Each nodeName As String In DBModifs.DBModifDefColl(control.Id).Keys
                Dim imageMsoStr As String = IIf(control.Id = "DBSeqnce", "ShowOnNewButton", IIf(control.Id = "DBMapper", "TableSave", IIf(control.Id = "DBAction", "TableIndexes", "undefined imageMso")))
                xmlString = xmlString + "<button id='_" + nodeName + "' label='do " + nodeName + "' imageMso='" + imageMsoStr + "' onAction='DBModifClick' tag='" + control.Id + "' screentip='display definitions for " + nodeName + "'/>"
            Next
            xmlString += "</menu>"
            Return xmlString
        Catch ex As Exception
            DBModifs.ErrorMsg("Exception caught while building xml: " + ex.Message)
            Return ""
        End Try
    End Function

    ''' <summary>show a screentip for the dynamic DBMapper/DBAction/DBSequence Menus (also showing the ID behind)</summary>
    ''' <param name="control"></param>
    ''' <returns>the screentip</returns>
    Public Function getDBModifScreentip(control As CustomUI.IRibbonControl) As String
        Return "Select DBModifier to store/do action/do sequence (" + control.Id + ")"
    End Function

    ''' <summary>to show the DBModif sheet button only if it was collected...</summary>
    ''' <param name="control"></param>
    ''' <returns>true if to be displayed</returns>
    Public Function getDBModifMenuVisible(control As CustomUI.IRibbonControl) As Boolean
        Try
            Return DBModifs.DBModifDefColl.ContainsKey(control.Id)
        Catch ex As Exception
            Return False
        End Try
    End Function

    ''' <summary>add DBModif</summary>
    ''' <param name="control"></param>
    Public Sub addDBModif(control As CustomUI.IRibbonControl)
        If ExcelDnaUtil.Application.ActiveWorkbook Is Nothing Then
            DBModifs.ErrorMsg("No active workbook to store CustomXMLParts into. Create one, please!")
        Else
            Dim DBModifType As String = InputBox("which DBModifType should be created (1..DBMapper, 2..DBAction or 3..DBSequence)?", "select DBModifType", "1")
            If DBModifType = vbNullString Then Exit Sub
            If CDbl(DBModifType) > 3 Or CDbl(DBModifType) < 1 Then
                DBModifs.ErrorMsg("no such DBModifType allowed !")
                Exit Sub
            End If
            DBModifType = Choose(CDbl(DBModifType), "DBMapper", "DBAction", "DBSeqnce")
            DBModifs.createDBModif(DBModifType)
            DBModifs.getDBModifDefinitions()
            theRibbon.Invalidate()
        End If
    End Sub

    ''' <summary>DBModif button activated, display the defined values in the DBModifier...</summary>
    ''' <param name="control"></param>
    Public Sub DBModifClick(control As CustomUI.IRibbonControl)
        Dim nodeName As String = Right(control.Id, Len(control.Id) - 1)
        DBModifs.DBModifDefColl(control.Tag).Item(nodeName).displayDBModif()
    End Sub

#Enable Warning IDE0060
End Class
