module MyFunctions

open ExcelDna.Integration

[<ExcelFunction(Description="My first .NET function")>]
let HelloDna name = 
    "Hello " + name