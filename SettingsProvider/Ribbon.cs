using ExcelDna.Integration.CustomUI;
using SettingsProvider.Properties;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace SettingsProvider
{
    [ComVisible(true)]
    public class Ribbon : ExcelRibbon
    {
        public override string GetCustomUI(string RibbonID)
        {
            return @"
      <customUI xmlns='http://schemas.microsoft.com/office/2006/01/customui'>
      <ribbon>
        <tabs>
          <tab id='tab1' label='My Tab'>
            <group id='group1' label='My Group'>
              <button id='buttonLoad' label='Load Settings' onAction='OnLoadSettingsPressed'/>
              <button id='buttonOverride' label='Override Settings' onAction='OnOverrideSettingsPressed'/>
              <button id='buttonSave' label='Save Settings' onAction='OnSaveSettingsPressed'/>
              <button id='buttonReset' label='Reset Settings' onAction='OnResetSettingsPressed'/>
            </group >
          </tab>
        </tabs>
      </ribbon>
    </customUI>";
        }

        public void OnLoadSettingsPressed(IRibbonControl control)
        {
            var magicNumber = Settings.Default.MagicNumber;
            var userName = Settings.Default.UserName;
            MessageBox.Show($"Magic Number:  {magicNumber}, User Name: {userName}");
        }

        public void OnOverrideSettingsPressed(IRibbonControl control)
        {
            //Settings.Default.AppKey = "EvenMoreMagix";
            Settings.Default.MagicNumber = 123.456;
            Settings.Default.UserName = "The real slim shady";
        }

        public void OnSaveSettingsPressed(IRibbonControl control)
        {
            Settings.Default.Save();
        }

        public void OnResetSettingsPressed(IRibbonControl control)
        {
            Settings.Default.Reset();
        }
    }
}
