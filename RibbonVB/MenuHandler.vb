Imports ExcelDna.Integration
Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop
Imports Microsoft.Vbe.Interop
Imports Microsoft.Office.Interop.Excel

''' <summary>handles all Menu related aspects (context menu for building/refreshing, "DBAddin"/"Load Config" tree menu for retrieving stored configuration files, etc.)</summary>
<ComVisible(True)>
Public Class MenuHandler
    Inherits CustomUI.ExcelRibbon

    ''' <summary>callback after Excel loaded the Ribbon, used to initialize data for the Ribbon</summary>
    Public Sub ribbonLoaded(theRibbon As CustomUI.IRibbonUI)
        Globals.theRibbon = theRibbon
    End Sub

    ''' <summary>creates the Ribbon (only at startup). any changes to the ribbon can only be done via dynamic menus</summary>
    ''' <returns></returns>
    Public Overrides Function GetCustomUI(RibbonID As String) As String
        ' Ribbon definition XML
        Dim customUIXml As String = "<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui' onLoad='ribbonLoaded' ><ribbon><tabs><tab id='RibbonAddinTab' label='RibbonVB'>"
        ' DBAddin Group: environment choice, DBConfics selection tree, purge names tool button and dialogBoxLauncher for AboutBox
        customUIXml +=
        "<group id='AddinGroup' label='Addin settings'>" +
            "<dropDown id='envDropDown' label='Environment:' sizeString='1234567890123456' getEnabled='GetEnabledSelect' getSelectedItemIndex='GetSelectedEnvironment' getItemCount='GetItemCount' getItemID='GetItemID' getItemLabel='GetItemLabel' getSupertip='getSelectedTooltip' onAction='selectEnvironment'/>" +
            "<buttonGroup id='buttonGroup1'>" +
                "<menu id='configMenu' label='Settings'>" +
                    "<button id='user' label='User settings' onAction='showAddinConfig' imageMso='ControlProperties' screentip='Show/edit user settings' />" +
                    "<button id='central' label='Central settings' onAction='showAddinConfig' imageMso='TablePropertiesDialog' screentip='Show/edit central settings' />" +
                    "<button id='addin' label='DBAddin settings' onAction='showAddinConfig' imageMso='ServerProperties' screentip='Show/edit standard Addin settings' />" +
                "</menu>" +
                "<button id='props' label='Workbook Properties' onAction='showCProps' getImage='getCPropsImage' screentip='Change custom properties:' getSupertip='getToggleCPropsScreentip' />" +
            "</buttonGroup>" +
        "</group>"
        ' Multiple Tools Group:
        customUIXml +=
        "<group id ='grp1' autoScale='false' centerVertically='false' label='Multiple Tools' getVisible='GetVisible' tag='grp1'>" +
            "<splitButton id='sbt3' getVisible='GetVisible' getEnabled='GetEnabled'>" +
                "<button id='btn_sbt3' imageMso='AlignLeft' label='SplitButton' tag='SplitButton' onAction='OnActionButton'/>" +
                "<menu id='mnu_sbt3' tag='mnu_sbt3'>" +
                    "<button id='btn_1' imageMso='AlignLeft' label='button1' tag='SplitButton_sub1' onAction='OnActionButton'/>" +
                    "<button id='btn_2' imageMso='AlignLeft' label='button2' tag='SplitButton_sub2' onAction='OnActionButton'/>" +
                "</menu>" +
            "</splitButton>" +
            "<checkBox id='cbx0' label='Checkbox' onAction='OnActionCheckbox' getPressed='GetPressedCheckbox' tag='DefaultValue:=0' getVisible='GetVisible' getEnabled='GetEnabled'/>" +
            "<separator id='sep12' getVisible='GetVisible'/>" +
            "<editBox id='ebx5' label='Editbox' getText='GetTextEditBox' onChange='OnChangeEditBox' tag='Editbox5' getVisible='GetVisible' getEnabled='GetEnabled'/>" +
            "<comboBox id='cmb7' label='ComboBox' onChange='OnChangeCombobox' getItemCount = 'GetItemCountCbx' getItemLabel = 'GetItemLabelCbx' getVisible='GetVisible' getEnabled='GetEnabled' tag='ComboBox7'/>" +
            "<toggleButton id='tgb6' size='large' label='ToggleButton' tag='ToggleButton6' onAction='OnActionTglButton' getPressed='GetPressedTglButton' getVisible='GetVisible' getEnabled='GetEnabled'/>" +
            "<gallery id='gal8' size='large' showItemLabel='true' label='Gallery' tag='Gallery8' getItemCount='GetItemCountGallery' getItemLabel='GetItemLabelGallery' onAction='OnActionGallery' getVisible='GetVisible' getEnabled='GetEnabled' rows='100' columns='3'/>" +
        "</group>"
        ' Tools Group:
        customUIXml +=
        "<group id='AddinToolsGroup' label='Addin Tools'>" +
            "<buttonGroup id='buttonGroup2'>" +
                "<dynamicMenu id='DBConfigs' label='Configs' imageMso='QueryShowTable' screentip='Configuration Files quick access' getContent='getDBConfigMenu'/>" +
            "</buttonGroup>" +
            "<buttonGroup id='buttonGroup3'>" +
                "<button id='cmdbutcreate' label='Cmdbutton' screentip='creates a commandbar button' imageMso='BorderErase' onAction='clickcmdbutcreate'/>" +
                "<button id='showLog' label='Log' screentip='shows Database Addins Diagnostic Display' getImage='getLogsImage' onAction='clickShowLog'/>" +
                "<button id='designmode' label='Buttons' onAction='showToggleDesignMode' getImage='getToggleDesignImage' getScreentip='getToggleDesignScreentip'/>" +
            "</buttonGroup>" +
        "</group>"
        ' DBModif Group: maximum three DBModif types possible (depending on existence in current workbook): 
        customUIXml +=
        "<group id='DBModifGroup' label='Execute DBModifier'>"
        For Each DBModifType As String In {"DBSeqnce", "DBMapper", "DBAction"}
            customUIXml += "<dynamicMenu id='" + DBModifType + "' " +
                                                "size='large' getLabel='getDBModifTypeLabel' imageMso='ApplicationOptionsDialog' " +
                                                "getScreentip='getDBModifScreentip' getContent='getDBModifMenuContent' getVisible='getDBModifMenuVisible'/>"
        Next
        customUIXml += "</group>"
        customUIXml += "</tab></tabs></ribbon>"
        ' Context menus for refresh, jump and creation: in cell, row, column and ListRange (area of ListObjects)
        customUIXml += "<contextMenus>" +
        "<contextMenu idMso ='ContextMenuCell'>" +
            "<button id='refreshDataC' label='button 1 (Ctl-Sh-R)' imageMso='Refresh' onAction='clickbutton1' insertBeforeMso='Cut'/>" +
            "<button id='gotoDBFuncC' label='button 2 (Ctl-Sh-J)' imageMso='ConvertTextToTable' onAction='clickbutton2' insertBeforeMso='Cut'/>" +
             "<menu id='createMenu' label='submenu' insertBeforeMso='Cut'>" +
                "<button id='DBMapperC' tag='1' label='button 1' imageMso='TableSave' onAction='clickCreateButton'/>" +
                "<button id='DBActionC' tag='2' label='button 2' imageMso='TableIndexes' onAction='clickCreateButton'/>" +
                "<button id='DBSequenceC' tag='3' label='button 3' imageMso='ShowOnNewButton' onAction='clickCreateButton'/>" +
                "<menuSeparator id='separator' />" +
                "<button id='DBListFetchC' tag='4' label='button 4' imageMso='GroupLists' onAction='clickCreateButton'/>" +
                "<button id='DBRowFetchC' tag='5' label='button 5' imageMso='GroupRecords' onAction='clickCreateButton'/>" +
                "<button id='DBSetQueryPivotC' tag='6' label='button 6' imageMso='AddContentType' onAction='clickCreateButton'/>" +
                "<button id='DBSetQueryListObjectC' tag='7' label='button 7' imageMso='AddContentType' onAction='clickCreateButton'/>" +
            "</menu>" +
            "<menuSeparator id='MySeparatorC' insertBeforeMso='Cut'/>" +
        "</contextMenu>" +
        "<contextMenu idMso ='ContextMenuPivotTable'>" +
            "<button id='refreshDataPT' label='button 1 (Ctl-Sh-R)' imageMso='Refresh' onAction='clickbutton1' insertBeforeMso='Copy'/>" +
            "<button id='gotoDBFuncPT' label='button 2 (Ctl-Sh-J)' imageMso='ConvertTextToTable' onAction='clickbutton2' insertBeforeMso='Copy'/>" +
            "<menuSeparator id='MySeparatorPT' insertBeforeMso='Copy'/>" +
        "</contextMenu>" +
        "<contextMenu idMso ='ContextMenuCellLayout'>" +
            "<button id='refreshDataCL' label='button 1 (Ctl-Sh-R)' imageMso='Refresh' onAction='clickbutton1' insertBeforeMso='Cut'/>" +
            "<button id='gotoDBFuncCL' label='button 2 (Ctl-Sh-J)' imageMso='ConvertTextToTable' onAction='clickbutton2' insertBeforeMso='Cut'/>" +
            "<menu id='createMenuCL' label='Insert/Edit DBFunc/DBModif' insertBeforeMso='Cut'>" +
                "<button id='DBMapperCL' tag='1' label='button 1' imageMso='TableSave' onAction='clickCreateButton'/>" +
                "<button id='DBActionCL' tag='2' label='button 2' imageMso='TableIndexes' onAction='clickCreateButton'/>" +
                "<button id='DBSequenceCL' tag='3' label='button 3' imageMso='ShowOnNewButton' onAction='clickCreateButton'/>" +
                "<menuSeparator id='separatorCL' />" +
                "<button id='DBListFetchCL' tag='4' label='button 4' imageMso='GroupLists' onAction='clickCreateButton'/>" +
                "<button id='DBRowFetchCL' tag='5' label='button 5' imageMso='GroupRecords' onAction='clickCreateButton'/>" +
                "<button id='DBSetQueryPivotCL' tag='6' label='button 6' imageMso='AddContentType' onAction='clickCreateButton'/>" +
                "<button id='DBSetQueryListObjectCL' tag='7' label='button 7' imageMso='AddContentType' onAction='clickCreateButton'/>" +
            "</menu>" +
            "<menuSeparator id='MySeparatorCL' insertBeforeMso='Cut'/>" +
        "</contextMenu>" +
        "<contextMenu idMso='ContextMenuRow'>" +
            "<button id='refreshDataR' label='button 1 (Ctl-Sh-R)' imageMso='Refresh' onAction='clickbutton1' insertBeforeMso='Cut'/>" +
            "<menuSeparator id='MySeparatorR' insertBeforeMso='Cut'/>" +
        "</contextMenu>" +
        "<contextMenu idMso='ContextMenuColumn'>" +
            "<button id='refreshDataZ' label='button 1 (Ctl-Sh-R)' imageMso='Refresh' onAction='clickbutton1' insertBeforeMso='Cut'/>" +
            "<menuSeparator id='MySeparatorZ' insertBeforeMso='Cut'/>" +
        "</contextMenu>" +
        "<contextMenu idMso='ContextMenuListRange'>" +
            "<button id='refreshDataL' label='button 1 (Ctl-Sh-R)' imageMso='Refresh' onAction='clickbutton1' insertBeforeMso='Cut'/>" +
            "<button id='gotoDBFuncL' label='button 2 (Ctl-Sh-J)' imageMso='ConvertTextToTable' onAction='clickbutton2' insertBeforeMso='Cut'/>" +
            "<menu id='createMenuL' label='submenu' insertBeforeMso='Cut'>" +
                "<button id='DBMapperL' tag='1' label='button 1' imageMso='TableSave' onAction='clickCreateButton'/>" +
                "<button id='DBSequenceL' tag='2' label='button 2' imageMso='ShowOnNewButton' onAction='clickCreateButton'/>" +
            "</menu>" +
            "<menuSeparator id='MySeparatorL' insertBeforeMso='Cut'/>" +
        "</contextMenu>" +
        "</contextMenus></customUI>"
        Return customUIXml
    End Function
    '<control idMso="FontColorPicker" imageMso="FontColorPicker"/>

#Disable Warning IDE0060 ' Hide not used Parameter warning as this is very often the case with the below callbacks from the ribbon

    Public Function GetVisible(control As CustomUI.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled(control As CustomUI.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetPressedCheckbox(control As CustomUI.IRibbonControl) As Boolean
        Return CheckboxState
    End Function

    Private CheckboxState As Boolean = False
    Public Sub OnActionCheckbox(control As CustomUI.IRibbonControl, pressed As Boolean)
        CheckboxState = pressed
    End Sub

    Public Sub OnActionButton(control As CustomUI.IRibbonControl)
        MsgBox("pressed " + control.Tag + ", checkbox state: " + CheckboxState.ToString())
    End Sub

    Private EditBoxText As String = "TestText"

    Public Function GetTextEditBox(control As CustomUI.IRibbonControl) As String
        Return EditBoxText
    End Function

    Public Sub OnChangeEditBox(control As CustomUI.IRibbonControl, strText As String)
        EditBoxText = strText
        MsgBox("entered " + strText + " in EditBox")
    End Sub

    Public Sub OnChangeCombobox(control As CustomUI.IRibbonControl, strText As String)
        If Not environdefs.Contains(strText) Then
            ReDim Preserve environdefs(environdefs.Length)
            environdefs(environdefs.Length - 1) = strText
            theRibbon.InvalidateControl("cmb7")
        Else
            MsgBox("existing entry " + strText + "!")
        End If
    End Sub

    Public Function GetItemCountCbx(control As CustomUI.IRibbonControl) As Integer
        Return Globals.environdefs.Length
    End Function

    Public Function GetItemLabelCbx(control As CustomUI.IRibbonControl, index As Integer) As String
        Return Globals.environdefs(index)
    End Function

    Public Sub OnActionTglButton(control As CustomUI.IRibbonControl, pressed As Boolean)
        Globals.DontChangeEnvironment = Not Globals.DontChangeEnvironment
        theRibbon.InvalidateControl("envDropDown")
    End Sub

    Public Function GetPressedTglButton(control As CustomUI.IRibbonControl) As Boolean
        Return Globals.DontChangeEnvironment
    End Function

    Public Sub OnActionGallery(control As CustomUI.IRibbonControl, id As String, index As Integer)
        MsgBox("Action Gallery: clicked on " + DisplayedNames(index))
    End Sub

    Private DisplayedNames As String() = {"Huey", "Dewey", "Louie", "Donald", "Scrooge", "Daisy", "Goofy", "Mickey", "Gus Goose", "Minnie", "Pluto"}
    Public Function GetItemCountGallery(ByVal Control As CustomUI.IRibbonControl) As Integer
        Return DisplayedNames.Length
    End Function

    Public Function GetItemLabelGallery(ByVal Control As CustomUI.IRibbonControl, index As Integer) As String
        Return DisplayedNames(index)
    End Function

    Public Function GetItemImageGallery(Control As CustomUI.IRibbonControl, index As Integer) As IPicture

    End Function

    ' display warning button icon on Cprops change if DBFskip is set
    Public Function getCPropsImage(control As CustomUI.IRibbonControl) As String
        If Not ExcelDnaUtil.Application.ActiveWorkbook Is Nothing Then
            For Each docproperty In ExcelDnaUtil.Application.ActiveWorkbook.CustomDocumentProperties
                If LCase(docproperty.Name) = "dbfskip" Then
                    If docproperty.Value Then Return "DeclineTask"
                End If
            Next
        End If
        Return "AcceptTask"
    End Function

    'display warning icon on log button if warning has been logged
    Public Function getLogsImage(control As CustomUI.IRibbonControl) As String
        If Globals.WarningIssued Then
            Return "IndexUpdate"
        Else
            Return "MailMergeStartLetters"
        End If
    End Function

    ' display state of designmode in screentip of dialogBox launcher
    ' returns screentip and the state of designmode
    Public Function getToggleCPropsScreentip(control As CustomUI.IRibbonControl) As String
        getToggleCPropsScreentip = ""
        If Not ExcelDnaUtil.Application.ActiveWorkbook Is Nothing Then
            Try
                Dim docproperty As Microsoft.Office.Core.DocumentProperty
                For Each docproperty In ExcelDnaUtil.Application.ActiveWorkbook.CustomDocumentProperties
                    getToggleCPropsScreentip += docproperty.Name + ":" + docproperty.Value.ToString + vbCrLf
                Next
            Catch ex As Exception
                getToggleCPropsScreentip += "exception when collecting docproperties: " + ex.Message
            End Try
        End If
    End Function

    ' click on change props: show builtin properties dialog
    Public Sub showCProps(control As CustomUI.IRibbonControl)
        If Not ExcelDnaUtil.Application.ActiveWorkbook Is Nothing Then
            ExcelDnaUtil.Application.Dialogs(Excel.XlBuiltInDialog.xlDialogProperties).Show
            ' to check whether DBFskip has changed:
            Globals.theRibbon.InvalidateControl(control.Id)
        End If
    End Sub

    ' toggle designmode button
    Public Sub showToggleDesignMode(control As CustomUI.IRibbonControl)
        Dim cbrs As Object = ExcelDnaUtil.Application.CommandBars
        If Not cbrs Is Nothing AndAlso cbrs.GetEnabledMso("DesignMode") Then
            cbrs.ExecuteMso("DesignMode")
        Else
            Globals.ErrorMsg("Couldn't toggle designmode, because Designmode commandbar button is not available (no button?)", "toggle Designmode", MsgBoxStyle.Exclamation)
        End If
        ' update state of designmode in screentip
        Globals.theRibbon.InvalidateControl(control.Id)
    End Sub

    ' display state of designmode in screentip of button
    ' returns screentip and the state of designmode
    Public Function getToggleDesignScreentip(control As CustomUI.IRibbonControl) As String
        Dim cbrs As Object = ExcelDnaUtil.Application.CommandBars
        If Not cbrs Is Nothing AndAlso cbrs.GetEnabledMso("DesignMode") Then
            Return "Designmode is currently " + IIf(cbrs.GetPressedMso("DesignMode"), "on !", "off !")
        Else
            Return "Designmode commandbar button not available (no button on sheet)"
        End If
    End Function

    ' display state of designmode in icon of button
    ' returns screentip and the state of designmode
    Public Function getToggleDesignImage(control As CustomUI.IRibbonControl) As String
        Dim cbrs As Object = ExcelDnaUtil.Application.CommandBars
        If Not cbrs Is Nothing AndAlso cbrs.GetEnabledMso("DesignMode") Then
            If cbrs.GetPressedMso("DesignMode") Then
                Return "ObjectsGroupMenuOutlook"
            Else
                Return "SelectMenuAccess"
            End If
        Else
            Return "SelectMenuAccess"
        End If
    End Function

    ' for environment dropdown to get the total number of the entries
    Public Function GetItemCount(control As CustomUI.IRibbonControl) As Integer
        Return Globals.environdefs.Length
    End Function

    ' for environment dropdown to get the label of the entries
    Public Function GetItemLabel(control As CustomUI.IRibbonControl, index As Integer) As String
        Return Globals.environdefs(index)
    End Function

    ' for environment dropdown to get the ID of the entries
    Public Function GetItemID(control As CustomUI.IRibbonControl, index As Integer) As String
        Return Globals.environdefs(index)
    End Function

    ' after selection of environment (using selectEnvironment) used to return the selected environment
    Public Function GetSelectedEnvironment(control As CustomUI.IRibbonControl) As Integer
        Return Globals.selectedEnvironment
    End Function

    ' tooltip for the environment select drop down
    ' returns the tooltip
    Public Function getSelectedTooltip(control As CustomUI.IRibbonControl) As String
        If Globals.DontChangeEnvironment Then
            Return "DontChangeEnvironment is set, therefore changing the Environment is prevented !"
        Else
            Return "Configurations in Addin config %appdata%\Microsoft\Addins\RibonAddin.xll.configs"
        End If
    End Function

    ' whether to enable environment select drop down
    ' true if enabled
    Public Function GetEnabledSelect(control As CustomUI.IRibbonControl) As Integer
        Return Not Globals.DontChangeEnvironment
    End Function

    ' Choose environment
    Public Sub selectEnvironment(control As CustomUI.IRibbonControl, id As String, index As Integer)
        Globals.selectedEnvironment = index
        MsgBox("selected Environment " + index.ToString())
    End Sub

    ' show xll standard config (AppSetting), central config (referenced by App Settings file attr) or user config (referenced by CustomSettings configSource attr)
    Public Sub showAddinConfig(control As CustomUI.IRibbonControl)
        MsgBox("AddinConfig: " + control.Tag)
    End Sub

    ' on demand, refresh the DB Config tree
    Public Sub refreshDBConfigTree(control As CustomUI.IRibbonControl)
        Globals.initSettings()
        ConfigFiles.createConfigTreeMenu()
        Globals.ErrorMsg("refreshed Config Tree Menu", "refresh Config tree...", MsgBoxStyle.Information)
        Globals.theRibbon.Invalidate()
    End Sub

    ' get DB Config Menu from File
    Public Function getDBConfigMenu(control As CustomUI.IRibbonControl) As String
        If ConfigFiles.ConfigMenuXML = vbNullString Then ConfigFiles.createConfigTreeMenu()
        Return ConfigFiles.ConfigMenuXML
    End Function

    ' load config if config tree menu end-button has been activated (path to config xcl file is in control.Tag)
    Public Sub getConfig(control As CustomUI.IRibbonControl)
        ConfigFiles.loadConfig(control.Tag)
    End Sub

    ' set the name of the DBModifType dropdown to the DB Modifier name
    Public Function getDBModifTypeLabel(control As CustomUI.IRibbonControl) As String
        getDBModifTypeLabel = If(control.Id = "DBSeqnce", "DBSequence", control.Id)
    End Function

    ' create the buttons in the DBModif sheet dropdown menu
    ' returns the menu content xml
    Public Function getDBModifMenuContent(control As CustomUI.IRibbonControl) As String
        Dim xmlString As String = "<menu xmlns='http://schemas.microsoft.com/office/2009/07/customui'>"
        Try
            If Not Globals.DBModifDefColl.ContainsKey(control.Id) Then Return ""
            Dim DBModifTypeName As String = IIf(control.Id = "DBSeqnce", "DBSequence", IIf(control.Id = "DBMapper", "DB Mapper", IIf(control.Id = "DBAction", "DB Action", "undefined DBModifTypeName")))
            For Each nodeName As String In Globals.DBModifDefColl(control.Id).Keys
                Dim descName As String = IIf(nodeName = control.Id, "Unnamed " + DBModifTypeName, Replace(nodeName, DBModifTypeName, ""))
                Dim imageMsoStr As String = IIf(control.Id = "DBSeqnce", "ShowOnNewButton", IIf(control.Id = "DBMapper", "TableSave", IIf(control.Id = "DBAction", "TableIndexes", "undefined imageMso")))
                Dim superTipStr As String = IIf(control.Id = "DBSeqnce", "executes " + DBModifTypeName + " defined in: " + nodeName, IIf(control.Id = "DBMapper", "stores data defined in DBMapper (named " + nodeName + ") range on " + Globals.DBModifDefColl(control.Id).Item(nodeName), IIf(control.Id = "DBAction", "executes Action defined in DBAction (named " + nodeName + ") range on " + Globals.DBModifDefColl(control.Id).Item(nodeName), "undefined superTip")))
                xmlString = xmlString + "<button id='_" + nodeName + "' label='do " + descName + "' imageMso='" + imageMsoStr + "' onAction='DBModifClick' tag='" + control.Id + "' screentip='do " + DBModifTypeName + ": " + descName + "' supertip='" + superTipStr + "' />"
            Next
            xmlString += "</menu>"
            Return xmlString
        Catch ex As Exception
            Globals.ErrorMsg("Exception caught while building xml: " + ex.Message)
            Return ""
        End Try
    End Function

    ' show a screentip for the dynamic DBMapper/DBAction/DBSequence Menus (also showing the ID behind)
    ' returns the screentip
    Public Function getDBModifScreentip(control As CustomUI.IRibbonControl) As String
        Return "Select DBModifier to store/do action/do sequence (" + control.Id + ")"
    End Function

    ' to show the DBModif sheet button only if it was collected
    ' returns true if to be displayed
    Public Function getDBModifMenuVisible(control As CustomUI.IRibbonControl) As Boolean
        Try
            Return Globals.DBModifDefColl.ContainsKey(control.Id)
        Catch ex As Exception
            Return False
        End Try
    End Function

    ' DBModif button activated, do DB Mapper/DB Action/DB Sequence or define existing (CtrlKey pressed)
    Public Sub DBModifClick(control As CustomUI.IRibbonControl)
        Dim nodeName As String = Right(control.Id, Len(control.Id) - 1)
        ' nice trick to check whether cell edit mode is active...
        If Not ExcelDnaUtil.Application.CommandBars.GetEnabledMso("FileNewDefault") Then
            Globals.ErrorMsg("Cannot execute DB Modifier while cell editing active !", "DB Modifier execution", MsgBoxStyle.Exclamation)
            Exit Sub
        End If
        Try
            If My.Computer.Keyboard.CtrlKeyDown And My.Computer.Keyboard.ShiftKeyDown Then
                MsgBox("editing DBModif " + nodeName + " ...")
            Else
                MsgBox("DBModif " + nodeName + " activated...")
            End If
        Catch ex As Exception
            Globals.ErrorMsg("Exception: " + ex.Message + ",control.Tag:" + control.Tag + ",nodeName:" + nodeName, "DBModif Click")
        End Try
    End Sub

    ' context menu button 1
    Public Sub clickbutton1(control As CustomUI.IRibbonControl)
        button1()
    End Sub

    ' context menu button 2
    Public Sub clickbutton2(control As CustomUI.IRibbonControl)
        button2()
    End Sub

    ' create cmd buttons in worksheet
    Public Sub clickcmdbutcreate(control As CustomUI.IRibbonControl)
        If Not ExcelDnaUtil.Application.ActiveWorkbook Is Nothing Then
            Dim cbshp As Excel.OLEObject = ExcelDnaUtil.Application.ActiveSheet.OLEObjects.Add(ClassType:="Forms.CommandButton.1", Link:=False, DisplayAsIcon:=False, Left:=600, Top:=70, Width:=120, Height:=24)
            Dim cb As Forms.CommandButton = cbshp.Object
            Dim cbName As String = InputBox("Name of DBModifier (should start with DBMapper, DBAction or DBSeqnce):", "create command button", "DBMapper")
            Try
                cb.Name = IIf(cbName = "", "DBMapper", cbName)
                cb.Caption = IIf(cbName = "", "Unnamed DBMapper", cbName)
            Catch ex As Exception
                cbshp.Delete()
                Globals.ErrorMsg("Couldn't name CommandButton '" + cbName + "': " + ex.Message, "CommandButton create Error")
                Exit Sub
            End Try
            If Len(cbName) > 31 Then
                cbshp.Delete()
                Globals.ErrorMsg("CommandButton codenames cannot be longer than 31 characters ! '" + cbName + "': ", "CommandButton create Error")
                Exit Sub
            End If
            ' fail to assign a handler? remove commandbutton (otherwise it gets hard to edit an existing DBModification with a different name).
            If Not AddInEvents.assignHandler(ExcelDnaUtil.Application.ActiveSheet) Then
                cbshp.Delete()
            End If
        End If
    End Sub

    ' show the trace log
    Public Sub clickShowLog(control As CustomUI.IRibbonControl)
        ExcelDna.Logging.LogDisplay.Show()
        ' reset warning flag
        WarningIssued = False
        theRibbon.InvalidateControl("showLog")
    End Sub

    ' context menu entries in sub menu
    Public Sub clickCreateButton(control As CustomUI.IRibbonControl)
        MsgBox("sub menu button " + control.Tag + " clicked...")
    End Sub

#Enable Warning IDE0060
End Class
