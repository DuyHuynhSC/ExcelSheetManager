using System.Drawing;
using System.Windows.Forms;
using System.Windows.Forms.Integration;
using ExcelSheetManager.ViewModels;

namespace ExcelSheetManager.Views
{
    public class TaskPaneHost : System.Windows.Forms.UserControl
    {
        private readonly ElementHost _elementHost;
        private readonly TaskPaneView _wpfView;

        public TaskPaneHost()
        {
            _elementHost = new ElementHost
            {
                Dock = DockStyle.Fill,
                BackColor = Color.FromArgb(15, 23, 42) // Match WPF background #0F172A
            };

            _wpfView = new TaskPaneView();
            _elementHost.Child = _wpfView;

            Controls.Add(_elementHost);
        }

        public MainViewModel? ViewModel => _wpfView.ViewModel;
    }
}
