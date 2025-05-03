using ExcelDna.Integration.CustomUI;
using System.Runtime.InteropServices;

namespace Ribbon
{
    [ComVisible(true)]
    public class RibbonController : ExcelRibbon
    {
        public override string GetCustomUI(string RibbonID)
        {
            return @"
      <customUI xmlns='http://schemas.microsoft.com/office/2006/01/customui'>
      <ribbon>
        <tabs>
          <tab id='tab1' label='My Tab'>
            <group id='group1' label='My Group'>
              <button id='button1' label='My Button' onAction='OnButtonPressed'/>
            </group >
          </tab>
        </tabs>
      </ribbon>
    </customUI>";
        }

        public void OnButtonPressed(IRibbonControl control)
        {
            new Form1 { Visible = true, TopMost = true };
        }
    }
}
