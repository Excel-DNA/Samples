using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;

namespace MasterSlave
{
    [ComVisible(true)]
    public class MasterRibbon : ExcelRibbon
    {
        public MasterRibbon()
        {
        }

        public override string GetCustomUI(string RibbonID)
        {
            // XlCall.Excel(XlCall.xlcMessage, "BANG!!!!");
            Console.Beep();
            Console.Beep();
            Console.Beep();
            return
@"
            <customUI xmlns = 'http://schemas.microsoft.com/office/2006/01/customui' >
                <ribbon>
                    <tabs>
                        <tab id = 'MyTab' label = 'Master' >
                            <group id = 'MyGroup' label = 'Slave Driver' >
                                <button id = 'RegisterSlave' label = 'Register Slave' onAction = 'RunTagMacro' tag = 'RegisterSlave' />
                                <button id = 'UnregisterSlave' label = 'Unregister Slave' onAction = 'RunTagMacro' tag = 'UnregisterSlave' />
                            </group>
                        </tab>
                    </tabs>
               </ribbon>
            </customUI>
";
        }
    }

}
