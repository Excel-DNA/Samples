using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;

namespace GenericHost
{
    public class MyCustomRibbonRegistration : IRibbonRegistration
    {
        public List<ExcelComAddIn> GetRibbonAddIns() => new List<ExcelComAddIn> { new MyCustomRibbon() };
    }
}
