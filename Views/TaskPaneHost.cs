using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Windows.Forms.Integration;
using ExcelSheetManager.ViewModels;

namespace ExcelSheetManager.Views
{
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    [ProgId("ExcelSheetManager.TaskPaneHost")]
    [Guid("A5F19D34-9B05-4B82-94C3-7E4A8D9183C2")]
    public class TaskPaneHost : System.Windows.Forms.UserControl
    {
        private ElementHost? _elementHost;
        private TaskPaneView? _wpfView;

        public TaskPaneHost()
        {
            // Keep constructor empty and lightweight for ActiveX CoCreateInstance in Office 365/2024
            this.BackColor = Color.FromArgb(15, 23, 42);
        }

        protected override void OnHandleCreated(EventArgs e)
        {
            base.OnHandleCreated(e);
            InitializeWpfChild();
        }

        private void InitializeWpfChild()
        {
            if (_elementHost == null)
            {
                try
                {
                    _elementHost = new ElementHost
                    {
                        Dock = DockStyle.Fill,
                        BackColor = Color.FromArgb(15, 23, 42)
                    };

                    _wpfView = new TaskPaneView();
                    _elementHost.Child = _wpfView;
                    Controls.Add(_elementHost);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"TaskPane WPF Load Error: {ex.Message}\n\n{ex.StackTrace}", "Excel Sheet Manager", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        public MainViewModel? ViewModel => _wpfView?.ViewModel;
    }
}
