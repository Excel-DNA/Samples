using System.Runtime.InteropServices;
using ExcelDna.Integration.CustomUI;

namespace CustomTaskPane
{
    [ComVisible(true)]
    public class MyRibbon : ExcelRibbon
    {
        public override string GetCustomUI(string RibbonID)
        {
            return
@"<customUI xmlns='http://schemas.microsoft.com/office/2006/01/customui' loadImage='LoadImage'>
    <ribbon>
    <tabs>
        <tab id='CustomTab' label='Custom Task Pane Test'>
        <group id='SampleGroup' label='CTP Control'>
            <button id='Button1' label='Show CTP' size='large' onAction='OnShowCTP' />
            <button id='Button2' label='Delete CTP' size='large' onAction='OnDeleteCTP' />
        </group >
        </tab>
    </tabs>
    </ribbon>
</customUI>
";
        }

        public void OnShowCTP(IRibbonControl control)
        {
            CTPManager.ShowCTP();
        }


        public void OnDeleteCTP(IRibbonControl control)
        {
            CTPManager.DeleteCTP();
        }
    }
}
