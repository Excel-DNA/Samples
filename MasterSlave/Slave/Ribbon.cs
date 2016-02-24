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
    public class Ribbon : ExcelRibbon
    {
        public Ribbon()
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
                        <tab id = 'CustomTab' label = 'My Tab' >
                            <group id = 'SampleGroup' label = 'My Sample Group' >
                                <button id = 'Button1' label = 'My Button Label' size = 'large' onAction = 'RunTagMacro' tag = 'ShowHelloMessage' />
                                <button id = 'Button2' label = 'My Second Button' size = 'normal' onAction = 'OnButtonPressed' />
                            </group>
                        </tab>
                    </tabs>
               </ribbon>
            </customUI>
";
        }
    }

}
