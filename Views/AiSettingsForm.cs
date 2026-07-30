using System;
using System.Drawing;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelSheetManager.Helpers;

namespace ExcelSheetManager.Views
{
    public class AiSettingsForm : Form
    {
        private TextBox _txtBaseUrl = null!;
        private TextBox _txtApiKey = null!;
        private TextBox _txtModelName = null!;
        private Button _btnTest = null!;
        private Button _btnSave = null!;
        private Label _lblStatus = null!;

        public AiSettingsForm()
        {
            this.Size = new Size(500, 330);
            this.Text = "⚙️ Local AI Settings (OpenAI Compatible)";
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

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

            int y = 16;

            // Base URL Label & Input
            Label lblUrl = new Label { Text = "OpenAI Base URL (Server Endpoint):", Location = new Point(20, y), AutoSize = true, Font = new Font("Segoe UI", 9f, FontStyle.Bold) };
            _txtBaseUrl = new TextBox { Location = new Point(20, y + 22), Width = 440 };
            y += 60;

            // Model Name Label & Input
            Label lblModel = new Label { Text = "Model Name (e.g. llama3, qwen2.5-coder, mistral):", Location = new Point(20, y), AutoSize = true, Font = new Font("Segoe UI", 9f, FontStyle.Bold) };
            _txtModelName = new TextBox { Location = new Point(20, y + 22), Width = 440 };
            y += 60;

            // API Key Label & Input
            Label lblKey = new Label { Text = "API Key (Optional / Default 'local'):", Location = new Point(20, y), AutoSize = true, Font = new Font("Segoe UI", 9f, FontStyle.Bold) };
            _txtApiKey = new TextBox { Location = new Point(20, y + 22), Width = 440 };
            y += 65;

            // Status Label
            _lblStatus = new Label { Text = "Enter local AI endpoint details above.", Location = new Point(20, y), AutoSize = true, ForeColor = Color.FromArgb(148, 163, 184) };

            // Buttons
            _btnTest = new Button
            {
                Text = "Test Connection",
                Location = new Point(220, y + 25),
                Width = 120,
                Height = 32,
                BackColor = Color.FromArgb(59, 130, 246),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            _btnTest.FlatAppearance.BorderSize = 0;
            _btnTest.Click += async (s, e) => await TestConnectionAsync();

            _btnSave = new Button
            {
                Text = "Save & Close",
                Location = new Point(350, y + 25),
                Width = 110,
                Height = 32,
                BackColor = Color.FromArgb(16, 185, 129),
                ForeColor = Color.White,
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
            this.Controls.Add(_lblStatus);
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
            _lblStatus.Text = "Testing connection to Local AI...";
            _lblStatus.ForeColor = Color.FromArgb(59, 130, 246);
            _btnTest.Enabled = false;

            try
            {
                bool ok = await AiService.TestConnectionAsync();
                if (ok)
                {
                    _lblStatus.Text = "✓ Connection successful! Local AI is online.";
                    _lblStatus.ForeColor = Color.FromArgb(16, 185, 129); // Green
                }
                else
                {
                    _lblStatus.Text = "✗ Could not connect. Verify Base URL & Model.";
                    _lblStatus.ForeColor = Color.FromArgb(239, 68, 68); // Red
                }
            }
            catch (Exception ex)
            {
                _lblStatus.Text = $"✗ Error: {ex.Message}";
                _lblStatus.ForeColor = Color.FromArgb(239, 68, 68);
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
