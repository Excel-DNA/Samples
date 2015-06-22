using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelDna.Integration;
using log4net;
using log4net.Config;

[assembly: log4net.Config.XmlConfigurator()]
namespace Logging
{

    // Used to test that a particular logging configuration is configured correctly
    public class TestLoggers : IExcelAddIn
    {
        public void AutoOpen()
        {
            // log4net
            // Configuration is done with an attribute: [assembly:log4net.Config.XmlConfigurator()]
            // Alternative is to call 
            // XmlConfigurator.Configure();
            // but for our example it would be too late.
            ILog log = LogManager.GetLogger(typeof(TestLoggers));   // Typically 
            log.Info("Testing log4net Info message");


            // NLog

        }

        public void AutoClose() { }
    }
}
