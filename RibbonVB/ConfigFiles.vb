Imports ExcelDna.Integration
Imports Microsoft.Office.Interop
Imports System.IO ' for getting config files for menu

'''<summary>procedures used for loading config files (containing DBFunctions and general sheet content) and building the config menu</summary>
Public Module ConfigFiles

    ''' <summary>loads config from file given in theFileName</summary>
    ''' <param name="theFileName">the File name of the config file</param>
    Public Sub loadConfig(theFileName As String)
        Dim ItemLine As String
        Dim retval As Integer

        retval = QuestionMsg("Inserting contents configured in " + theFileName, MsgBoxStyle.OkCancel, "DBAddin: Inserting Configuration...", MsgBoxStyle.Information)
        If retval = vbCancel Then Exit Sub
        If ExcelDnaUtil.Application.ActiveWorkbook Is Nothing Then ExcelDnaUtil.Application.Workbooks.Add

        ' open file for reading
        Try
            Dim fileReader As System.IO.StreamReader = My.Computer.FileSystem.OpenTextFileReader(theFileName, Text.Encoding.Default)
            Do
                ItemLine = fileReader.ReadLine()
                ' ConfigArray: Configs are tab separated pairs of <RC location vbTab function formula> vbTab <...> vbTab...
                Dim ConfigArray As String() = Split(ItemLine, vbTab)
                createFunctionsInCells(ExcelDnaUtil.Application.ActiveCell, ConfigArray)
            Loop Until fileReader.EndOfStream
            fileReader.Close()
        Catch ex As Exception
            Globals.ErrorMsg("Error (" + ex.Message + ") during filling items from config file '" + theFileName + "' in ConfigFiles.loadConfig")
        End Try
    End Sub

    ''' <summary>creates functions in target cells (relative to referenceCell) as defined in ItemLineDef</summary>
    ''' <param name="originCell">original reference Cell</param>
    ''' <param name="ItemLineDef">String array, pairwise containing relative cell addresses and the functions in those cells (= cell content)</param>
    Public Sub createFunctionsInCells(originCell As Excel.Range, ByRef ItemLineDef As Object)
        Dim cellToBeStoredAddress As String, cellToBeStoredContent As String
        ' disabling calculation is necessary to avoid object errors
        Dim calcMode As Long = ExcelDnaUtil.Application.Calculation
        ExcelDnaUtil.Application.Calculation = Excel.XlCalculation.xlCalculationManual
        Dim i As Long

        ' for each defined cell address and content pair
        For i = 0 To UBound(ItemLineDef) Step 2
            cellToBeStoredAddress = ItemLineDef(i)
            cellToBeStoredContent = ItemLineDef(i + 1)

            ' get cell in relation to function target cell
            If cellToBeStoredAddress.Length > 0 Then
                ' if there is a reference to a different sheet in cellToBeStoredAddress (starts with '<sheetname>'! ) and this sheet doesn't exist, create it...
                If InStr(1, cellToBeStoredAddress, "!") > 0 Then
                    Dim theSheetName As String = Replace(Mid$(cellToBeStoredAddress, 1, InStr(1, cellToBeStoredAddress, "!") - 1), "'", "")
                    Try
                        Dim testSheetExist As String = ExcelDnaUtil.Application.Worksheets(theSheetName).name
                    Catch ex As Exception
                        With ExcelDnaUtil.Application.Worksheets.Add(After:=originCell.Parent)
                            .name = theSheetName
                        End With
                        originCell.Parent.Activate()
                    End Try
                End If

                ' get target cell respecting relative cellToBeStoredAddress starting from originCell
                Dim TargetCell As Excel.Range = Nothing
                If Not getRangeFromRelative(originCell, cellToBeStoredAddress, TargetCell) Then
                    Globals.ErrorMsg("Excel Borders would be violated by placing target cell (relative address:" + cellToBeStoredAddress + ")" + vbLf + "Cell content: " + cellToBeStoredContent + vbLf + "Please select different cell !!")
                End If

                ' finally fill function target cell with function text (relative cell references to target cell) or value
                Try
                    If Left$(cellToBeStoredContent, 1) = "=" Then
                        TargetCell.FormulaR1C1 = cellToBeStoredContent
                    Else
                        TargetCell.Value = cellToBeStoredContent
                    End If
                Catch ex As Exception
                    Globals.ErrorMsg("Error in setting Cell: " + ex.Message, "Create functions in cells")
                End Try
            End If
        Next
        ExcelDnaUtil.Application.Calculation = calcMode
    End Sub

    ''' <summary>gets target range in relation to origin range</summary>
    ''' <param name="originCell">the origin cell to be related to</param>
    ''' <param name="relAddress">the relative address of the target as an RC style reference</param>
    ''' <param name="theTargetRange">the returned resulting range</param>
    ''' <returns>True if boundaries are not violated, false otherwise</returns>
    Private Function getRangeFromRelative(originCell As Excel.Range, ByVal relAddress As String, ByRef theTargetRange As Excel.Range) As Boolean
        Dim theSheetName As String

        If InStr(1, relAddress, "!") = 0 Then
            theSheetName = originCell.Parent.Name
        Else
            theSheetName = Replace(Mid$(relAddress, 1, InStr(1, relAddress, "!") - 1), "'", "")
        End If
        ' parse row or column out of RC style reference adresses
        Dim startRow As Long = 0, startCol As Long = 0, endRow As Long = 0, endCol As Long = 0
        Dim begins As String
        Dim relAddressPart() As String = Split(relAddress, ":")

        ' get startRow and startCol from both multi and single cell range (without separation by ":")
        If InStr(1, relAddressPart(0), "R[") > 0 Then
            begins = Mid$(relAddressPart(0), InStr(1, relAddressPart(0), "R[") + 2)
            startRow = CLng(Mid$(begins, 1, InStr(1, begins, "]") - 1))
        End If
        If InStr(1, relAddressPart(0), "C[") > 0 Then
            begins = Mid$(relAddressPart(0), InStr(1, relAddressPart(0), "C[") + 2)
            startCol = CLng(Mid$(begins, 1, InStr(1, begins, "]") - 1))
        End If
        ' get endRow and endCol in case of multi cell range ((topleftAddress):(bottomrightAddress))
        If UBound(relAddressPart) = 1 Then
            If InStr(1, relAddressPart(1), "R[") > 0 Then
                begins = Mid$(relAddressPart(1), InStr(1, relAddressPart(1), "R[") + 2)
                endRow = CLng(Mid$(begins, 1, InStr(1, begins, "]") - 1))
            End If
            If InStr(1, relAddressPart(1), "C[") > 0 Then
                begins = Mid$(relAddressPart(1), InStr(1, relAddressPart(1), "C[") + 2)
                endCol = CLng(Mid$(begins, 1, InStr(1, begins, "]") - 1))
            End If
        End If
        ' check if resulting target range would violate excel sheets boundaries, if so, then return error (false)
        If originCell.Row + startRow > 0 And originCell.Row + startRow <= originCell.Parent.Rows.Count _
           And originCell.Column + startCol > 0 And originCell.Column + startCol <= originCell.Parent.Columns.Count Then
            If InStr(1, relAddress, ":") > 0 Then
                ' for multi cell relative ranges, final target offset is starting at the bottom right of relative range
                theTargetRange = ExcelDnaUtil.Application.Range(originCell, originCell.Offset(endRow - startRow, endCol - startCol))
            Else
                ' for single cell relative ranges, target range is just set to the offsetting row and column of the relative range.
                theTargetRange = originCell
            End If
            theTargetRange = ExcelDnaUtil.Application.Worksheets(theSheetName).Range(theTargetRange.Offset(startRow, startCol).Address)
            getRangeFromRelative = True
        Else
            theTargetRange = Nothing
            getRangeFromRelative = False
        End If
    End Function

    ''' <summary>the folder used to store predefined DB item definitions</summary>
    Public ConfigStoreFolder As String
    ''' <summary>Array of special ConfigStoreFolders for non default treatment of Name Separation (Camelcase) and max depth</summary>
    Public specialConfigStoreFolders() As String
    ''' <summary>fixed max Depth for Ribbon</summary>
    Const maxMenuDepth As Integer = 5
    ''' <summary>fixed max size for menu XML</summary>
    Const maxSizeRibbonMenu = 320000
    ''' <summary>used to create menu and button ids</summary>
    Private menuID As Integer
    ''' <summary>tree menu stored here</summary>
    Public ConfigMenuXML As String = vbNullString
    ''' <summary>individual limitation of grouping of entries in special folders (set by _DBname_MaxDepth)</summary>
    Public specialFolderMaxDepth As Integer
    ''' <summary>store found submenus in this collection</summary>
    Private specialConfigFoldersTempColl As Collection
    ''' <summary>for correct display of menu</summary>
    Private ReadOnly xnspace As XNamespace = "http://schemas.microsoft.com/office/2009/07/customui"

    ''' <summary>creates the Config tree menu by reading the menu elements from the config store folder files/subfolders</summary>
    Public Sub createConfigTreeMenu()
        Dim currentBar, button As XElement

        For Each tModule As ProcessModule In Process.GetCurrentProcess().Modules
            Dim sModule As String = tModule.FileName
            If sModule.ToUpper.Contains("RIBBONVB") Then
                ConfigStoreFolder = Left(tModule.FileName, InStrRev(tModule.FileName, "\",) - 1) + "\..\..\"
                Exit For
            End If
        Next

        If ExcelDnaUtil.Application.ActiveWorkbook Is Nothing Then
            Globals.ErrorMsg("No active workbook, setting ConfigStoreFolder to " + ConfigStoreFolder + " !")
        ElseIf ExcelDnaUtil.Application.ActiveWorkbook.Path = "" Then
            Globals.ErrorMsg("No path for active workbook, setting ConfigStoreFolder to " + ConfigStoreFolder + " !")
        Else
            ConfigStoreFolder = ExcelDnaUtil.Application.ActiveWorkbook.Path
        End If

        ' top level menu
        currentBar = New XElement(xnspace + "menu")
        ' add refresh button to top level
        button = New XElement(xnspace + "button")
        button.SetAttributeValue("id", "refreshConfig")
        button.SetAttributeValue("label", "refresh DBConfig Tree")
        button.SetAttributeValue("imageMso", "Refresh")
        button.SetAttributeValue("onAction", "refreshDBConfigTree")
        currentBar.Add(button)
        ' collect all config files recursively, creating submenus for the structure (see readAllFiles) and buttons for the final config files.
        specialConfigFoldersTempColl = New Collection
        menuID = 0
        readAllFiles(ConfigStoreFolder, currentBar)
        specialConfigFoldersTempColl = Nothing
        ExcelDnaUtil.Application.StatusBar = ""
        currentBar.SetAttributeValue("xmlns", xnspace)
        ' avoid exception in ribbon...
        ConfigMenuXML = currentBar.ToString()
        If ConfigMenuXML.Length > maxSizeRibbonMenu Then
            MsgBox("Too many entries in " + ConfigStoreFolder + ", can't display them in a ribbon menu ..")
            ConfigMenuXML = "<menu xmlns='" + xnspace.ToString() + "'><button id='refreshDBConfig' label='refresh DBConfig Tree' imageMso='Refresh' onAction='refreshDBConfigTree'/></menu>"
        End If
    End Sub

    ''' <summary>reads all files contained in rootPath and its subfolders (recursively) and adds them to the DBConfig menu (sub)structure (recursively). For folders contained in specialConfigStoreFolders, apply further structuring by splitting names on camelcase or specialConfigStoreSeparator</summary>
    ''' <param name="rootPath">root folder to be searched for config files</param>
    ''' <param name="currentBar">current menu element, where submenus and buttons are added</param>
    ''' <param name="Folderpath">for sub menus path of current folder is passed (recursively)</param>
    Private Sub readAllFiles(rootPath As String, ByRef currentBar As XElement, Optional Folderpath As String = vbNullString)
        Try
            Dim newBar As XElement = Nothing
            Static MenuFolderDepth As Integer = 1 ' needed to not exceed max. menu depth (currently 5)

            ' read all leaf node entries (files) and sort them by name to create action menus
            Dim di As DirectoryInfo = New DirectoryInfo(rootPath)
            Dim fileList() As FileSystemInfo = di.GetFileSystemInfos("*.xcl").OrderBy(Function(fi) fi.Name).ToArray()
            If fileList.Length > 0 Then
                For i = 0 To UBound(fileList)
                    newBar = New XElement(xnspace + "button")
                    menuID += 1
                    newBar.SetAttributeValue("id", "m" + menuID.ToString())
                    newBar.SetAttributeValue("screentip", "click to insert DBListFetch for " + Left$(fileList(i).Name, Len(fileList(i).Name) - 4) + " in active cell")
                    newBar.SetAttributeValue("tag", rootPath + "\" + fileList(i).Name)
                    newBar.SetAttributeValue("label", Folderpath + Left$(fileList(i).Name, Len(fileList(i).Name) - 4))
                    newBar.SetAttributeValue("onAction", "getConfig")
                    currentBar.Add(newBar)
                Next
            End If

            ' read all folder xcl entries and sort them by name
            Dim DirList() As DirectoryInfo = di.GetDirectories().OrderBy(Function(fi) fi.Name).ToArray()
            If DirList.Length = 0 Then Exit Sub
            ' recursively build branched menu structure from dirEntries
            For i = 0 To UBound(DirList)
                ExcelDnaUtil.Application.StatusBar = "Filling DBConfigs Menu: " + rootPath + "\" + DirList(i).Name
                ' only add new menu element if below max. menu depth for ribbons
                If MenuFolderDepth < maxMenuDepth Then
                    newBar = New XElement(xnspace + "menu")
                    menuID += 1
                    newBar.SetAttributeValue("id", "m" + menuID.ToString())
                    newBar.SetAttributeValue("label", DirList(i).Name)
                    currentBar.Add(newBar)
                    MenuFolderDepth += 1
                    readAllFiles(rootPath + "\" + DirList(i).Name, newBar, Folderpath + DirList(i).Name + "\")
                    MenuFolderDepth -= 1
                Else
                    newBar = currentBar
                    readAllFiles(rootPath + "\" + DirList(i).Name, newBar, Folderpath + DirList(i).Name + "\")
                End If
            Next
        Catch ex As Exception
            Globals.ErrorMsg("Error (" + ex.Message + ") in MenuHandler.readAllFiles")
        End Try
    End Sub

    ''' <summary>parses Substrings (filenames in special Folders) contained in nameParts (recursively) of passed xcl config filepath (fullPathName) and adds them to currentBar and submenus (recursively)</summary>
    ''' <param name="nameParts">tokenized string (separated by space)</param>
    ''' <param name="currentBar">current menu element, where submenus and buttons are added</param>
    ''' <param name="fullPathName">full path name to xcl config file</param>
    ''' <param name="newRootName">the new root name for the menu, used avoid multiple placement of buttons in submenus</param>
    ''' <param name="Folderpath">Path of enclosing Folder(s)</param>
    ''' <param name="MenuFolderDepth">required for keeping maxMenuDepth limit</param>
    ''' <returns>new bar as Xelement (for containment)</returns>
    Private Function buildFileSepMenuCtrl(nameParts As String, ByRef currentBar As XElement, fullPathName As String, newRootName As String, Folderpath As String, MenuFolderDepth As Integer) As XElement
        Static MenuDepth As Integer = 0
        Try
            Dim newBar As XElement
            ' end node: add callable entry (= button)
            If InStr(1, nameParts, " ") = 0 Or MenuDepth + MenuFolderDepth >= maxMenuDepth Then
                Dim entryName As String = Mid$(fullPathName, InStrRev(fullPathName, "\") + 1)
                newBar = New XElement(xnspace + "button")
                menuID += 1
                newBar.SetAttributeValue("id", "m" + menuID.ToString())
                newBar.SetAttributeValue("screentip", "click to insert DBListFetch for " + Left$(entryName, Len(entryName) - 4) + " in active cell")
                newBar.SetAttributeValue("label", Left$(entryName, Len(entryName) - 4))
                newBar.SetAttributeValue("tag", fullPathName)
                newBar.SetAttributeValue("onAction", "getConfig")
                currentBar.Add(newBar)
                buildFileSepMenuCtrl = newBar
            Else  ' branch node: add new menu, recursively descend
                Dim newName As String = Left$(nameParts, InStr(1, nameParts, " ") - 1)
                ' prefix already exists: put new submenu below already existing prefix
                If specialConfigFoldersTempColl.Contains(newRootName + newName) Then
                    newBar = specialConfigFoldersTempColl(newRootName + newName)
                Else
                    newBar = New XElement(xnspace + "menu")
                    menuID += 1
                    newBar.SetAttributeValue("id", "m" + menuID.ToString())
                    newBar.SetAttributeValue("label", newName)
                    specialConfigFoldersTempColl.Add(newBar, newRootName + newName)
                    currentBar.Add(newBar)
                End If
                MenuDepth += 1
                buildFileSepMenuCtrl(Mid$(nameParts, InStr(1, nameParts, " ") + 1), newBar, fullPathName, newRootName + newName, Folderpath, MenuFolderDepth)
                MenuDepth -= 1
                buildFileSepMenuCtrl = newBar
            End If
        Catch ex As Exception
            Globals.ErrorMsg("Error (" + ex.Message + ") in MenuHandler.buildFileSepMenuCtrl")
            buildFileSepMenuCtrl = Nothing
        End Try
    End Function

End Module