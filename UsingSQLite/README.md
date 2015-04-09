# Using SQLite with Excel-DNA

This sample shows how to use the SQLite database library from an Excel-DNA add-in.

To try it:

## Create the Northwind.db database

* The sample database needs to be created. Download the file Northwind.Sqlite3.sql from https://code.google.com/p/northwindextended/downloads/detail?name=Northwind.Sqlite3.sql (or from a GitHub export like https://github.com/djhenderson/northwindextended/blob/master/Northwind.Sqlite3.sql)
* Read this .sql into a new database called Northwind.db, perhaps using the sqlite3 command-line utility:

  * sqlite.exe C:\Temp\Northwind.db
  * sqlite> BEGIN TRANSACTION;
  * sqlite> .read C:/Temp/Northwind.Sqlite3.sql
  * sqlite> COMMIT TRANSACTION;

## Create the Excel-DNA add-in

* Create a new C# Class Library project
* PM> Install-Package Excel-DNA
* PM> Install-Package System.Data.SQLite.Core
* Add the code in MyFunctions.cs
* Press F5 to run
* Enter =ProductName(1) in a cell - the result should be "Chai"
