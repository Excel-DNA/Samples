* Create new C# project of type "Class Library (.NET Framework)" called DnaComServer

* Package Manager Console:
	`PM> Install-Package ExcelDna.AddIn`

* Rename Class1.cs to AddIn.cs with this code:

```c#
	using System.Runtime.InteropServices;
	using ExcelDna.ComInterop;
	using ExcelDna.Integration;

	namespace DnaComServer
	{
		[ComVisible(true)]
		[ClassInterface(ClassInterfaceType.AutoDual)]
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

* DnaComServer-AddIn.dna file:

```xml
	<DnaLibrary Name="DnaComServer Add-In" RuntimeVersion="v4.0">
	  <ExternalLibrary Path="DnaComServer.dll" ExplicitExports="false" ComServer="true" LoadFromBytes="true" Pack="true" />
	</DnaLibrary>
```
* Fix Project -> Debug settings to remove %1 in "Start external program" path if present

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

* Set up the tlbexp run in the post-build step

```
	REM Setting up environment vairables
	call "$(DevEnvDir)..\..\VC\Auxiliary\Build\vcvarsall.bat" x86

	REM Temporarily copy ExcelDna.Integration.dll into output
	REM Note: Might need to change depending on where packages directory is
	copy "$(ProjectDir)\packages\ExcelDna.Integration.0.34.6\lib\ExcelDna.Integration.dll" "$(TargetDir)"

	REM Create .tlb file
	tlbexp.exe "$(ProjectDir)$(OutDir)$(TargetName)$(TargetExt)" /out:"$(ProjectDir)$(OutDir)$(TargetName).tlb"

	REM Delete extra copy of ExcelDna.Integration.dll from output
	del "$(TargetDir)ExcelDna.Integration.dll"

	REM Re-run the packing to include the .tlb inside the packed files for distribution
	REM Note: Might need to change depending on where packages directory is
	"$(ProjectDir)\packages\ExcelDna.AddIn.0.34.6\tools\ExcelDnaPack.exe" "$(ProjectDir)$(OutDir)$(TargetName)-AddIn.dna" /Y  /O "$(ProjectDir)$(OutDir)$(TargetName)-AddIn-packed.xll"
	"$(ProjectDir)\packages\ExcelDna.AddIn.0.34.6\tools\ExcelDnaPack.exe" "$(ProjectDir)$(OutDir)$(TargetName)-AddIn64.dna" /Y  /O "$(ProjectDir)$(OutDir)$(TargetName)-AddIn64-packed.xll"

	REM Register COM servers in add-in on this machine for testing
	REM Note: Change this to -AddIn64.xll if the 64-bit version of Excel is installed
	regsvr32.exe /s "$(ProjectDir)$(OutDir)$(TargetName)-AddIn.xll"
```

* F5 to build and run Excel

* (If this does not work due to tlbexp.exe not found, see [this discussion](https://groups.google.com/forum/#!topic/exceldna/XH3UbPwCnak).)

* Open VBA IDE (Alt+F11)

* Insert a new Module

* Add a reference to the .tlb:
	* Tools -> References
	* Find and tick DnaComServer in the list

* Put this code in to test late-bound COM Server

```vb
	Sub TestEarlyBound()
		Dim dnaComServer As DnaComServer.ComLibrary
		Dim hello As String
		Dim result As Double
    
		Set dnaComServer = New DnaComServer.ComLibrary
		hello = dnaComServer.ComLibraryHello()
		result = dnaComServer.Add(1, 2)
    
		Debug.Print hello, result
	End Sub
```
* Put cursor inside the Sub and press F5 to test - check Immediate window for output

