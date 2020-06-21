using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;

namespace GenericHost
{
    internal class ComAPI
    {
        public const string gstrIRibbonExtensibility = "000C0396-0000-0000-C000-000000000046";
    }

    [ComImport]
    [Guid(ComAPI.gstrIRibbonExtensibility)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    internal interface IRibbonExtensibility
    {
        [DispId(1)]
        string GetCustomUI(string RibbonID);
    }

    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public class MyCustomRibbon : ExcelComAddIn, IRibbonExtensibility
    {
        public string GetCustomUI(string RibbonID) => RibbonResources.CustomUI;

        public void OnButtonPressed(IRibbonControl control)
        {
            MessageBox.Show("Hello from control " + control.Id);
        }
    }
}
