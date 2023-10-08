# Using R.NET in Excel

This is just a basic sample to get started using R with Excel.

I followed these steps to create the add-in:

1. Ensure that R is installed from https://cran.r-project.org/bin/windows/base/
   The bitness of the R installation must match that of Excel. Note that v4.1.3 is the last version to support 32bit. 

2. Create a new "Class Library" project in Visual Studio.

3. In the NuGet package Manager Console, execute the commands:

        PM> Install-Package ExcelDna.Addin
        PM> Install-Package R.NET

4. Add the sample code from AddIn.cs.

5. F5 to run the add-in in Excel.

6. Enter the formula =TestRDotNet()` and `=MyRNorm(5)`. Numbers appear in Excel.


## Other links

* The R project, including links to the R installation is at, http://www.r-project.org/ 

* The R.NET project is at https://github.com/rdotnet/rdotnet
