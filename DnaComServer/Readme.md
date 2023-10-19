These are work in progress instructions for supporting COM server with the current ExcelDNA version. Note that 32bit Excel is not supported for early binding yet.

* Create new C# project of type "Class Library" called DnaComServer

* Package Manager Console:
	`PM> Install-Package ExcelDna.AddIn`
	`PM> Install-Package dSPACE.Runtime.InteropServices.BuildTasks`

* Rename Class1.cs to AddIn.cs with this code:

```c#
	using System.Runtime.InteropServices;
using ExcelDna.ComInterop;
using ExcelDna.Integration;

namespace DnaComServer
{
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IComLibrary
    {
        string ComLibraryHello();
        double Add(double x, double y);
    }

    [ComDefaultInterface(typeof(IComLibrary))]
    public class ComLibrary
    {
        public string ComLibraryHello()
        {
            return "Hello from DnaComServer.ComLibrary";
        }

        public double Add(double x, double y)
        {
            return x + y;
        }
    }

    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IComLibrary2
    {
        string ComLibrary2Hello();
        double Add2(double x, double y);
    }

    [ComDefaultInterface(typeof(IComLibrary2))]
    public class ComLibrary2
    {
        public string ComLibrary2Hello()
        {
            return "Hello from DnaComServer.ComLibrary2";
        }

        public double Add2(double x, double y)
        {
            return x + y;
        }
    }

    [ComVisible(false)]
    public class ExcelAddin : IExcelAddIn
    {
        public void AutoOpen()
        {
            ComServer.DllRegisterServer();
        }
        public void AutoClose()
        {
            ComServer.DllUnregisterServer();
        }
    }

    public static class Functions
    {
        [ExcelFunction]
        public static object DnaComServerHello()
        {
            return "Hello from DnaComServer!";
        }
    }
}
```

* Edit the DnaComServer.csproj file and add the following:

```xml
	<Project Sdk="Microsoft.NET.Sdk">

	<PropertyGroup>
		<TargetFrameworks>net472;net6.0-windows</TargetFrameworks>
		<ExcelAddInComServer>true</ExcelAddInComServer>
	</PropertyGroup>

	<ItemGroup>
		<PackageReference Include="ExcelDna.Addin" Version="*-*" />
		<PackageReference Include="dSPACE.Runtime.InteropServices.BuildTasks" Version="1.2.0"/>	
	</ItemGroup>

</Project>
```

* Press F5 to build and load in Excel

* Check formula =DnaComServerHello() in a sheet

* Open VBA IDE (Alt+F11)

* Insert a new Module

* Put this code in to test late-bound COM Server

```vb
	Sub TestLateBound()
		Dim dnaComServer As Object
		Dim hello As String
		Dim result As Double
    
		Set dnaComServer = CreateObject("DnaComServer.ComLibrary")
		hello = dnaComServer.ComLibraryHello()
		result = dnaComServer.Add(1, 2)
    
		Debug.Print hello, result
	End Sub
```

* Run by clicking inside Sub and pressing F5

* Check Immediate window for output:
	Hello from DnaComServer.ComLibrary         3 

* F5 to build and run Excel

* Open VBA IDE (Alt+F11)

* Insert a new Module

* Add a reference to the .tlb:
	* Tools -> References
	* Find and tick DnaComServer in the list

* Put this code in to test late-bound COM Server

```vb
	Sub TestEarlyBound()
		Dim dnaComServer As DnaComServer.IComLibrary
		Dim hello As String
		Dim result As Double
    
		Set dnaComServer = New DnaComServer.ComLibrary
		hello = dnaComServer.ComLibraryHello()
		result = dnaComServer.Add(1, 2)
    
		Debug.Print hello, result
	End Sub
```
* Put cursor inside the Sub and press F5 to test - check Immediate window for output

