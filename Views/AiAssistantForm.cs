using System;
using System.Drawing;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelSheetManager.Helpers;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelSheetManager.Views
{
    public class AiAssistantForm : Form
    {
        private static AiAssistantForm? _currentInstance;

        public static void ShowForm()
        {
            try
            {
                if (_currentInstance != null && !_currentInstance.IsDisposed)
                {
                    _currentInstance.Activate();
                    return;
                }

                _currentInstance = new AiAssistantForm();
                _currentInstance.Show();
                _currentInstance.Activate();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Could not open AI Assistant: {ex.Message}", "AI Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private TextBox _txtPrompt = null!;
        private TextBox _txtResponse = null!;
        private ComboBox _cmbMode = null!;
        private Button _btnSend = null!;
        private Button _btnCopy = null!;
        private Button _btnInsertCell = null!;
        private Button _btnSettings = null!;
        private Label _lblStatus = null!;

        private bool _isDarkTheme = true;

        public AiAssistantForm()
        {
            this.Size = new Size(580, 480);
            this.Text = "🤖 Local AI Assistant for Excel";
            this.StartPosition = FormStartPosition.CenterScreen;
            this.TopMost = true;

            _isDarkTheme = SettingsHelper.GetIsDarkTheme();
            this.BackColor = _isDarkTheme ? Color.FromArgb(15, 23, 42) : Color.FromArgb(241, 245, 249);
            this.ForeColor = _isDarkTheme ? Color.FromArgb(248, 250, 252) : Color.FromArgb(15, 23, 42);
            this.Font = new Font("Segoe UI", 9.5f);

            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();

            // Top Panel
            Panel pnlTop = new Panel
            {
                Dock = DockStyle.Top,
                Height = 130,
                Padding = new Padding(12, 10, 12, 6)
            };

            Label lblMode = new Label { Text = "AI Task Mode:", Location = new Point(12, 10), AutoSize = true, Font = new Font("Segoe UI", 9f, FontStyle.Bold) };
            _cmbMode = new ComboBox
            {
                Location = new Point(110, 6),
                Width = 240,
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            _cmbMode.Items.Add("⚡ Generate Excel Formula");
            _cmbMode.Items.Add("🔍 Explain Sheet / Formula");
            _cmbMode.Items.Add("💬 General Excel Assistant");
            _cmbMode.SelectedIndex = 0;

            _btnSettings = new Button
            {
                Text = "⚙️ Settings",
                Location = new Point(440, 6),
                Width = 110,
                Height = 26,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            _btnSettings.Click += (s, e) => OpenSettings();

            _txtPrompt = new TextBox
            {
                Location = new Point(12, 40),
                Width = 420,
                Height = 75,
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                Text = "Write a formula to calculate total sum of B2:B50 if A2:A50 equals 'Paid'"
            };

            _btnSend = new Button
            {
                Text = "🚀 Send",
                Location = new Point(440, 40),
                Width = 110,
                Height = 75,
                BackColor = Color.FromArgb(59, 130, 246),
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 10f, FontStyle.Bold),
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            _btnSend.FlatAppearance.BorderSize = 0;
            _btnSend.Click += async (s, e) => await SendPromptAsync();

            pnlTop.Controls.Add(lblMode);
            pnlTop.Controls.Add(_cmbMode);
            pnlTop.Controls.Add(_btnSettings);
            pnlTop.Controls.Add(_txtPrompt);
            pnlTop.Controls.Add(_btnSend);

            // Bottom Panel
            Panel pnlBottom = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 44,
                Padding = new Padding(12, 6, 12, 6)
            };

            _lblStatus = new Label
            {
                Dock = DockStyle.Left,
                AutoSize = true,
                TextAlign = ContentAlignment.MiddleLeft,
                Text = "Ready to connect to Local AI model.",
                ForeColor = Color.FromArgb(148, 163, 184)
            };

            _btnCopy = new Button
            {
                Dock = DockStyle.Right,
                Width = 120,
                Text = "📋 Copy Text",
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            _btnCopy.Click += (s, e) => CopyResponse();

            _btnInsertCell = new Button
            {
                Dock = DockStyle.Right,
                Width = 150,
                Text = "📥 Insert into Cell",
                BackColor = Color.FromArgb(16, 185, 129),
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 9f, FontStyle.Bold),
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            _btnInsertCell.FlatAppearance.BorderSize = 0;
            _btnInsertCell.Click += (s, e) => InsertIntoActiveCell();

            pnlBottom.Controls.Add(_lblStatus);
            pnlBottom.Controls.Add(_btnInsertCell);
            pnlBottom.Controls.Add(_btnCopy);

            // Middle Response Panel
            _txtResponse = new TextBox
            {
                Dock = DockStyle.Fill,
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Vertical,
                Font = new Font("Consolas", 10f),
                BackColor = _isDarkTheme ? Color.FromArgb(30, 41, 59) : Color.White,
                ForeColor = _isDarkTheme ? Color.FromArgb(248, 250, 252) : Color.FromArgb(15, 23, 42)
            };

            this.Controls.Add(_txtResponse);
            this.Controls.Add(pnlTop);
            this.Controls.Add(pnlBottom);

            this.ResumeLayout(false);
        }

        private void OpenSettings()
        {
            using (AiSettingsForm form = new AiSettingsForm())
            {
                form.ShowDialog();
            }
        }

        private async Task SendPromptAsync()
        {
            string prompt = _txtPrompt.Text.Trim();
            if (string.IsNullOrEmpty(prompt)) return;

            _btnSend.Enabled = false;
            _lblStatus.Text = "Sending request to Local AI...";
            _lblStatus.ForeColor = Color.FromArgb(59, 130, 246);
            _txtResponse.Text = "Generating AI response...";

            string systemPrompt = _cmbMode.SelectedIndex switch
            {
                0 => "You are an expert Excel formula generator. Output ONLY the exact Excel formula starting with =. Do not include markdown code block formatting or explanation unless asked.",
                1 => "You are an Excel data analyst. Explain the user's Excel formula or data clearly and concisely.",
                _ => "You are an intelligent AI assistant integrated inside Microsoft Excel."
            };

            try
            {
                string result = await AiService.GetCompletionAsync(prompt, systemPrompt);
                _txtResponse.Text = result;
                _lblStatus.Text = "AI Response received successfully.";
                _lblStatus.ForeColor = Color.FromArgb(16, 185, 129);
            }
            catch (Exception ex)
            {
                _txtResponse.Text = $"Error: {ex.Message}";
                _lblStatus.Text = "Error communicating with AI.";
                _lblStatus.ForeColor = Color.FromArgb(239, 68, 68);
            }
            finally
            {
                _btnSend.Enabled = true;
            }
        }

        private void CopyResponse()
        {
            if (!string.IsNullOrEmpty(_txtResponse.Text))
            {
                Clipboard.SetText(_txtResponse.Text);
                _lblStatus.Text = "Copied AI response to clipboard.";
            }
        }

        private void InsertIntoActiveCell()
        {
            try
            {
                string text = _txtResponse.Text.Trim();
                if (string.IsNullOrEmpty(text)) return;

                // Strip markdown backticks if present
                if (text.StartsWith("```"))
                {
                    int firstLine = text.IndexOf('\n');
                    int lastLine = text.LastIndexOf("```");
                    if (firstLine >= 0 && lastLine > firstLine)
                    {
                        text = text.Substring(firstLine + 1, lastLine - firstLine - 1).Trim();
                    }
                }

                var excelApp = (Excel.Application)ExcelDnaUtil.Application;
                if (excelApp?.ActiveCell != null)
                {
                    excelApp.ActiveCell.Value2 = text;
                    _lblStatus.Text = $"Inserted into active cell {excelApp.ActiveCell.Address[false, false]}";
                }
            }
            catch (Exception ex)
            {
                _lblStatus.Text = $"Insert error: {ex.Message}";
            }
        }
    }
}
