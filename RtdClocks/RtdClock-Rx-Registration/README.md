This project has the following NuGet packages installed:
* ExcelDna.AddIn
* System.Reactive

The registration APIs are in the `ExcelDna.Registration` namespace, but since Excel-DNA v1.9 they are included in `ExcelDna.Integration`, which is supplied by the `ExcelDna.AddIn` package. This sample should not reference or pack a separate `ExcelDna.Registration` package or DLL.

Additional project properties
* Add <ExcelAddInExplicitRegistration>true</ExcelAddInExplicitRegistration>
