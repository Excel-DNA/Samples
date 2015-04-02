# Using DAO with Excel-DNA (VB.NET)

This is a simple VB.NET project that shows how to get started using DAO with an Excel-DNA add-in.

DAO is the best library for interacting with Access databases, even when using .NET.

The project was created by following these steps:

* Create a new "Class Library" VB.NET project.

* Open the Tools-> NuGet Package Manager -> Package Manager Console.

* Install the following three packages: 
```
    PM> Install-Package Excel-DNA 
    PM> Install-Package Excel-DNA.Interop 
    PM> Install-Package Excel-DNA.Interop.DAO
```    
    
* Put the code in the AddIn.vb file.
    
* Press F5 to run - then check the functions and the menu button under the Add-Ins tab.

One issue is if you want your add-in to work with the 64-bit version of Excel (not common, since the 32-bit version is the default install even if Windows itself is 64-bit). 
Then you'll have to switch to the AceDAO library instead of the regular JET one. All your code should stay the same, but you need to change references and users might need to install the Acess 2013 runtime from here: http://www.microsoft.com/en-us/download/details.aspx?id=39358.

For 32-bit versions of Excel, no additional libraries should be required to use the JET database engine and the DAO classes.
