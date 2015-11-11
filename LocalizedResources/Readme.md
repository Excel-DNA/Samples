## Localized Resources

This project shows how to create and access localized resources from inside an Excel-DNA add-in.

The project was created with the following steps:
* Create new Class Library

Install ExcelDna.AddIn package:
    PM> Install-Package ExcelDna.AddIn

Add a test function:
    public class Class1
    {
        public static string locHello()
        {
            return "Hello from LocalizedResources";
        }
	}

Add the resources:
* Add new "Resources File"
* In Resource1.rex, edit the String1 resource\
* Copy and paste the Resource1.resx file, and rename the new copy to "Resource1.fr-FR.resx".
* Edit the French String1 to be "Première chaîne".

* Copy and paste, rename to "Resource1.es.resx"
* Edit the Spanish String1 to be "Primera cadena"


Add an accessor function:
    public class Class1
    {
	    //...

        public static string locGetString1(string cultureName)
        {
            return Resource1.ResourceManager.GetString("String1", CultureInfo.GetCultureInfo(cultureName));
        }
    }

Add packing directives to the .dna file to include the localized resources:
	<DnaLibrary Name="LocalizedResources Add-In" RuntimeVersion="v4.0">
	  <ExternalLibrary Path="LocalizedResources.dll" LoadFromBytes="true" Pack="true" />
	  <Reference Path="fr-FR\LocalizedResources.resources" Pack="true" />
      <Reference Path="es\LocalizedResources.resources" Pack="true" />
	</DnaLibrary>

Test and run:
	* Press F5 to build and load in Excel
	* =locHello()
	* =locGetString1()
	* =locGetString1("en")
	* =locGetString1("fr-FR")	
	* =locGetString1("es-AR")  (Note fallback to "es" resources)

TODO:
	* Implement and test packing for resources - expected for Excel-DNA v. 0.34.
