using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;

namespace GenericHost
{
    public class AddIn : IExcelAddIn
    {
        public void AutoOpen()
        {
            var config = BuildConfiguration();
            Debug.Print("AutoOpen - Enter");
            Task.Run(async () =>
            {
                await Task.Delay(10000);
                Debug.Print("AutoOpen - Enqueue Configure");
                ExcelAsyncUtil.QueueAsMacro(() =>
                {
                    config.Register();
                    Debug.Print("AutoOpen - Register");
                });
                await Task.Delay(10000);
                Debug.Print("AutoOpen - Enqueue Configure 2");
                ExcelAsyncUtil.QueueAsMacro(() =>
                {
                    config.Register();
                    Debug.Print("AutoOpen - Register 2");
                });
            });
            //ExcelAsyncUtil.QueueAsMacro(() =>
            //{
            //    Debug.Print("AutoOpen - QueueAsMacro");
            //    BuildConfiguration()
            //        .Register();
            //    Debug.Print("AutoOpen - Register");
            //});
            Debug.Print("AutoOpen - Exit");

        }

        public void AutoClose()
        {
        }

        static AddInConfiguration BuildConfiguration()
        {
            return new AddInConfiguration
            {
                //FunctionRegistration = new MyCustomFunctionRegistration(),
                RibbonRegistration = new MyCustomRibbonRegistration()
            };
        }

        
        [ExcelCommand(ShortCut ="^R")]
        public static void RegisterRibbon()
        {
            var ribbon = new MyCustomRibbon();
            ExcelComAddInHelper.LoadComAddIn(ribbon);
        }
    }
}
