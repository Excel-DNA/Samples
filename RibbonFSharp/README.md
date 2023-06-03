# F# Ribbon example

This is a small project to show how the ExcelRibbon can be used from an F# add-in.
The interesting code is in MyRibbon.fs

## Creating the ribbon project
* Create a new "Library" F# project
* Install-Package Excel-DNA
* Fix up the Debug path, as explained in the Readme.txt
* Copy the test code into Library1.fs
* F5 to run and check that the funciton works in Excel
* Add new file with the code in MyRibbon.fs
* Add a reference to System.Windows.Forms.dll
* Install-Package Excel-DNA.Interop
* F5 to run and check that the ribbon is there and the buttons work

## Using the COM object model
For macros that manipulate Excel, using the COM object model is best:
* Add references to the Primary Interop Assemblies (Install-Package ExcelDna.Interop) and 
* get hold of the root Application object with a call to ExcelDnaUtil.Application