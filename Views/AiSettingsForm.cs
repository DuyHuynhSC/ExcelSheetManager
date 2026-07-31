using System;
using System.Drawing;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelSheetManager.Helpers;

namespace ExcelSheetManager.Views
{
    internal class AiSettingsForm : Form
    {
        private TextBox _txtBaseUrl = null!;
        private TextBox _txtApiKey = null!;
        private TextBox _txtModelName = null!;
        private TextBox _txtDiagnosticLog = null!;
        private Button _btnTest = null!;
        private Button _btnSave = null!;

        public AiSettingsForm()
        {
            this.Size = new Size(535, 410);
            this.Text = "Local AI Connection Settings";
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.TopMost = true;

            bool isDark = SettingsHelper.GetIsDarkTheme();
            this.BackColor = isDark ? Color.FromArgb(15, 23, 42) : Color.FromArgb(241, 245, 249);
            this.ForeColor = isDark ? Color.FromArgb(248, 250, 252) : Color.FromArgb(15, 23, 42);
            this.Font = new Font("Segoe UI", 9.5f);

            InitializeComponent();
            LoadSettings();
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();

            bool isDark = SettingsHelper.GetIsDarkTheme();
            int y = 14;

            // Base URL Label & Input
            Label lblUrl = new Label { Text = "OpenAI Base URL / Full Endpoint:", Location = new Point(20, y), AutoSize = true, Font = new Font("Segoe UI", 9f, FontStyle.Bold) };
            _txtBaseUrl = new TextBox { Location = new Point(20, y + 22), Width = 480 };
            y += 58;

            // Model Name Label & Input
            Label lblModel = new Label { Text = "Model Name (e.g. llama3, qwen2.5-coder, mistral):", Location = new Point(20, y), AutoSize = true, Font = new Font("Segoe UI", 9f, FontStyle.Bold) };
            _txtModelName = new TextBox { Location = new Point(20, y + 22), Width = 480 };
            y += 58;

            // API Key Label & Input
            Label lblKey = new Label { Text = "API Key (Optional / Default 'local'):", Location = new Point(20, y), AutoSize = true, Font = new Font("Segoe UI", 9f, FontStyle.Bold) };
            _txtApiKey = new TextBox { Location = new Point(20, y + 22), Width = 480 };
            y += 58;

            // Diagnostic Log Box
            Label lblDiag = new Label { Text = "Connection Diagnostic Log:", Location = new Point(20, y), AutoSize = true, Font = new Font("Segoe UI", 8.5f, FontStyle.Bold) };
            _txtDiagnosticLog = new TextBox
            {
                Location = new Point(20, y + 20),
                Width = 480,
                Height = 85,
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Vertical,
                Font = new Font("Consolas", 8.5f),
                BackColor = isDark ? Color.FromArgb(30, 41, 59) : Color.White,
                ForeColor = isDark ? Color.FromArgb(248, 250, 252) : Color.FromArgb(15, 23, 42),
                Text = "Click 'Test Connection' to verify server status and request URL."
            };
            y += 112;

            // Buttons
            _btnTest = new Button
            {
                Text = "Test Connection",
                Location = new Point(230, y),
                Width = 140,
                Height = 32,
                BackColor = Color.FromArgb(59, 130, 246),
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 9f, FontStyle.Bold),
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            _btnTest.FlatAppearance.BorderSize = 0;
            _btnTest.Click += async (s, e) => await TestConnectionAsync();

            _btnSave = new Button
            {
                Text = "Save & Close",
                Location = new Point(380, y),
                Width = 120,
                Height = 32,
                BackColor = Color.FromArgb(16, 185, 129),
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 9f, FontStyle.Bold),
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            _btnSave.FlatAppearance.BorderSize = 0;
            _btnSave.Click += (s, e) => SaveSettings();

            this.Controls.Add(lblUrl);
            this.Controls.Add(_txtBaseUrl);
            this.Controls.Add(lblModel);
            this.Controls.Add(_txtModelName);
            this.Controls.Add(lblKey);
            this.Controls.Add(_txtApiKey);
            this.Controls.Add(lblDiag);
            this.Controls.Add(_txtDiagnosticLog);
            this.Controls.Add(_btnTest);
            this.Controls.Add(_btnSave);

            this.ResumeLayout(false);
        }

        private void LoadSettings()
        {
            _txtBaseUrl.Text = SettingsHelper.GetAiBaseUrl();
            _txtModelName.Text = SettingsHelper.GetAiModelName();
            _txtApiKey.Text = SettingsHelper.GetAiApiKey();
        }

        private async Task TestConnectionAsync()
        {
            SaveValuesToHelper();
            _txtDiagnosticLog.Text = $"Testing connection...\nTarget URL: {AiService.GetEndpointUrl()}\nModel: {_txtModelName.Text}...";
            _btnTest.Enabled = false;

            try
            {
                var (success, logMessage) = await AiService.TestConnectionDiagnosticsAsync();
                _txtDiagnosticLog.Text = logMessage;
            }
            catch (Exception ex)
            {
                _txtDiagnosticLog.Text = $"✗ Error testing connection: {ex.Message}";
            }
            finally
            {
                _btnTest.Enabled = true;
            }
        }

        private void SaveValuesToHelper()
        {
            SettingsHelper.SetAiBaseUrl(_txtBaseUrl.Text.Trim());
            SettingsHelper.SetAiModelName(_txtModelName.Text.Trim());
            SettingsHelper.SetAiApiKey(_txtApiKey.Text.Trim());
        }

        private void SaveSettings()
        {
            SaveValuesToHelper();
            this.DialogResult = DialogResult.OK;
            this.Close();
        }
    }
}
