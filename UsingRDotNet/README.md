# Using R.NET in Excel

This is just a basic sample to get started using R with Excel.

I followed these steps to create the add-in:

1. Ensure that R is installed. In my Windows "Add or Remove Programs" list I see "R for Windows 3.02".

2. Create a new "Class Library" project in Visual Studio.

3. In the NuGet package Manager Console, execute the commands:

        PM> Install-Package Excel-DNA
        PM> Install-Package R.NET.Community

4. Add the sample code from AddIn.cs.

5. F5 to run the add-in in Excel.

6. Enter the formula =TestRDotNet()` and `=MyRNorm(5)`. Numbers appear in Excel.


## Other links

* The R project, including links to the R installation is at, http://www.r-project.org/.

* The R.NET project is at http://rdotnet.codeplex.com/.

* A nice and more useful introduction to using R from Excel, using F#, was written by Natallie Baikevich and can be found at http://type-nat.ch/excel-dna-three-stories//
