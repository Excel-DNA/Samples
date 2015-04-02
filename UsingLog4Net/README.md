# Using Log4Net in an Excel-DNA add-in

This sample contains a simple project with an AppDomain configuration file that is used to configure the log4net logging library.

To create the project I followed these steps:
    
* Create a new Class Library project, in this case called "UsingLog4Net".

* Open the NuGet package manager and add these packages:
```
    PM> Install-Package Excel-DNA
    PM> Install-Package log4net
```

* Add a new "Application Configuration File" item to the project, and set its name to "UsingLog4Net-AddIn.xll.config". The name must exactly match the name of the add-in we are building, by default this is the project name with "-AddIn" added.

* Set the properties on "UsingLog4Net-AddIn.xll.config" to *Copy to Output Directory: Copy if newer*. This will ensure the file is copied next to the .xll file.

* Set the text in the .xll.config file to:

```xml
<?xml version="1.0" encoding="utf-8" ?>
<configuration>
  <configSections>
    <section name="log4net" type="log4net.Config.Log4NetConfigurationSectionHandler, log4net" />
  </configSections>
  <appSettings>
    <!-- Change this setting to "true" if you want to debug the log4net configuration -->
    <add key="log4net.Internal.Debug" value="false"/>
  </appSettings>
  <log4net>
    <appender name="DebugAppender" type="log4net.Appender.DebugAppender" >
      <layout type="log4net.Layout.PatternLayout">
        <conversionPattern value="%date [%thread] %-5level %logger [%property{NDC}] - %message%newline" />
      </layout>
    </appender>
    <root>
      <level value="DEBUG" />
      <appender-ref ref="DebugAppender" />
    </root>
  </log4net>
</configuration>
```

* Set the test code in the .cs file:

```C#
using ExcelDna.Integration;
using log4net;

[assembly:log4net.Config.XmlConfigurator()]

namespace UsingLog4Net
{
    public static class MyAddIn
    {
        static ILog Logger = LogManager.GetLogger("MyAddIn");

        public static double AddThemAndLog(double x, double y)
        {
            Logger.Debug(">>>>> AddThemAndLog called.");
            return x + y;
        }
    }
}
```

* Press F5 to build and load the add-in in Excel. 

* Enter as the formula into a cell: `=AddThemAndLog(2,3)`

* Check in the Visual Studio Output window:
```
    MyAddIn: 2015-04-02 23:42:09,032 [1] DEBUG MyAddIn [(null)] - >>>>> AddThemAndLog called.
```

* Further details on log4net configuration with alternative 'Appenders' can be found at: http://logging.apache.org/log4net/
