using System.Windows.Forms;
using System.Runtime.InteropServices;

namespace CustomTaskPane
{
    public interface IMyUserControl { }
    
    [ComDefaultInterface(typeof(IMyUserControl))]
    public partial class ContentControl : UserControl, IMyUserControl
    {
        public ContentControl()
        {
            InitializeComponent();
        }
    }
}
