Imports System.IO
Imports Microsoft.Office.Interop.Access.Dao
Imports ExcelDna.Integration

Public Module MyAddIn
    ' This would not be needed in VBA - in .NET we must instantiate the root COM object
    Private MyDBEngine As New DBEngine()
    Private MyDatabase As Database

    Const MyDatabasePath As String = "TestDB.mdb"

    Private Sub EnsureDatabaseIsConnected()
        If MyDatabase Is Nothing Then

            ' We've not set the database yet - try to open or create a new one
            If File.Exists(MyDatabasePath) Then
                MyDatabase = MyDBEngine.OpenDatabase(MyDatabasePath)
            Else
                ' Create new DB for testing - this could be OpenDatabase too
                MyDatabase = MyDBEngine.CreateDatabase(MyDatabasePath, Locale:=LanguageConstants.dbLangGeneral)

                ' Add the tables or give an error ....
                ' ...
            End If
        End If
    End Sub

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' These will be UDF functions you can call from a sheet...
    ' Check that the DAO is working correctly
    <ExcelFunction(Description:="Returns the version of the loaded DAO DBEngine")>
    Public Function DBEngineVersion() As Object
        Return MyDBEngine.Version
    End Function

    ' Check that the database was opened 
    <ExcelFunction(Description:="Returns the name of the open database")>
    Public Function DatabaseName() As String
        EnsureDatabaseIsConnected()
        Return MyDatabase.Name
    End Function

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' This will be a macro added to the Add-Ins tab on the ribbon
    <ExcelCommand(MenuName:="Excel-DNA DAO Sample", MenuText:="Show Database Name")>
    Public Sub ShowDatabaseName()
        EnsureDatabaseIsConnected()
        MsgBox(MyDatabase.Name, Title:="DAO Sample")
    End Sub

End Module
