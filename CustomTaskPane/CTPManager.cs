using System.Windows.Forms;
using ExcelDna.Integration.CustomUI;

namespace CustomTaskPane
{
    internal static class CTPManager
    {
        // WARNING: This won't work well under Excel 2013. There you need a different policy, since a CTP is attached only to a single window (one workbook).
        //          So having a single variable here means you can only ever have one CTP in one of the Excel 2013 windows.
        //          Maybe have a map from workbook to CTP, or have a floating one or something...

        static ExcelDna.Integration.CustomUI.CustomTaskPane ctp;

        public static void ShowCTP()
        {
            if (ctp == null)
            {
                // Make a new one using ExcelDna.Integration.CustomUI.CustomTaskPaneFactory 
                ctp = CustomTaskPaneFactory.CreateCustomTaskPane(typeof(ContentControl), "My Super Task Pane");
                ctp.Visible = true;
                ctp.DockPosition = MsoCTPDockPosition.msoCTPDockPositionLeft;
                ctp.DockPositionStateChange += ctp_DockPositionStateChange;
                ctp.VisibleStateChange += ctp_VisibleStateChange;
            }
            else
            {
                // Just show it again
                ctp.Visible = true;
            }
        }


        public static void DeleteCTP()
        {
            if (ctp != null)
            {
                // Could hide instead, by calling ctp.Visible = false;
                ctp.Delete();
                ctp = null;
            }
        }

        static void ctp_VisibleStateChange(ExcelDna.Integration.CustomUI.CustomTaskPane CustomTaskPaneInst)
        {
            MessageBox.Show("Visibility changed to " + CustomTaskPaneInst.Visible);
        }

        static void ctp_DockPositionStateChange(ExcelDna.Integration.CustomUI.CustomTaskPane CustomTaskPaneInst)
        {
            ((ContentControl)ctp.ContentControl).label1.Text = "Moved to " + CustomTaskPaneInst.DockPosition.ToString();
        }
        
    }
}
