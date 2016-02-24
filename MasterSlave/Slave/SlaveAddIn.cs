using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;

namespace MasterSlave
{
    public class SlaveAddIn : IExcelAddIn
    {
        public void AutoOpen()
        {
            Console.Beep();
            Console.Beep();
        }

        public void AutoClose()
        {
            Console.Beep();
            Console.Beep();
            Console.Beep();
            Console.Beep();
            Console.Beep();
            Console.Beep();
        }
    }
}
