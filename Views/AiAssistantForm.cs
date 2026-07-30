using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelSheetManager.Helpers;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelSheetManager.Views
{
    internal class AiAssistantForm : Form
    {
        private static AiAssistantForm? _currentInstance;

        internal static void ShowForm()
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
            this.Size = new Size(580, 490);
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
                Height = 135,
                Padding = new Padding(12, 10, 12, 6)
            };

            Label lblMode = new Label { Text = "AI Task Mode:", Location = new Point(12, 10), AutoSize = true, Font = new Font("Segoe UI", 9f, FontStyle.Bold) };
            _cmbMode = new ComboBox
            {
                Location = new Point(110, 6),
                Width = 250,
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            _cmbMode.Items.Add("⚡ Generate Excel Formula");
            _cmbMode.Items.Add("🌐 Translate Cell / Range (Dịch thuật)");
            _cmbMode.Items.Add("🔍 Explain Sheet / Formula");
            _cmbMode.Items.Add("💬 General Excel Assistant");
            _cmbMode.SelectedIndex = 1; // Default to Translate mode

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
                Height = 80,
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                Text = "Hãy dịch nội dung tại cell B1 từ tiếng Nhật sang tiếng Việt"
            };

            _btnSend = new Button
            {
                Text = "🚀 Send",
                Location = new Point(440, 40),
                Width = 110,
                Height = 80,
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
                Text = "Specify cell (e.g. B1) in prompt for auto-reading Excel data.",
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
                form.TopMost = true;
                form.ShowDialog(this);
            }
        }

        private async Task SendPromptAsync()
        {
            string prompt = _txtPrompt.Text.Trim();
            if (string.IsNullOrEmpty(prompt)) return;

            _btnSend.Enabled = false;
            _lblStatus.Text = "Reading Excel cells & sending to Local AI...";
            _lblStatus.ForeColor = Color.FromArgb(59, 130, 246);
            _txtResponse.Text = "Reading Excel cells and generating AI response...";

            // Auto-extract cell content referenced in user prompt
            string enrichedPrompt = EnrichPromptWithExcelCellData(prompt);

            string systemPrompt = _cmbMode.SelectedIndex switch
            {
                0 => "You are an expert Excel formula generator. Output ONLY the exact Excel formula starting with =. Do not include markdown code block formatting or explanation unless asked.",
                1 => "You are an expert translator integrated in Microsoft Excel. Read the provided cell content carefully and translate it accurately as requested. Output ONLY the translated text without extra conversational fluff.",
                2 => "You are an Excel data analyst. Explain the user's Excel formula or data clearly and concisely.",
                _ => "You are an intelligent AI assistant integrated inside Microsoft Excel."
            };

            try
            {
                string result = await AiService.GetCompletionAsync(enrichedPrompt, systemPrompt);
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

        private string EnrichPromptWithExcelCellData(string userPrompt)
        {
            try
            {
                var excelApp = (Excel.Application)ExcelDnaUtil.Application;
                if (excelApp == null || excelApp.ActiveSheet == null) return userPrompt;

                // Match cell references like B1, A1:C10, Sheet1!B1
                var matches = System.Text.RegularExpressions.Regex.Matches(
                    userPrompt,
                    @"\b([A-Za-z]{1,3}[1-9][0-9]{0,6}(?::[A-Za-z]{1,3}[1-9][0-9]{0,6})?)\b",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase
                );

                List<string> foundRefs = new List<string>();
                foreach (System.Text.RegularExpressions.Match m in matches)
                {
                    string val = m.Value.Trim();
                    // Filter out common non-cell words
                    if (val.Length == 1 || (val.Length <= 3 && !val.Any(char.IsDigit))) continue;
                    if (!foundRefs.Contains(val, StringComparer.OrdinalIgnoreCase))
                    {
                        foundRefs.Add(val);
                    }
                }

                StringBuilder sb = new StringBuilder();
                sb.AppendLine(userPrompt);

                if (excelApp?.ActiveSheet is Excel.Worksheet activeWs)
                {
                    bool addedContext = false;

                    if (foundRefs.Count > 0)
                    {
                        sb.AppendLine("\n--- DỮ LIỆU ĐỌC TỰ ĐỘNG TỪ EXCEL ---");
                        foreach (string cellRef in foundRefs)
                        {
                            try
                            {
                                Excel.Range? range = activeWs.get_Range(cellRef);
                                if (range != null)
                                {
                                    object raw = range.Value2;
                                    string textVal = raw != null ? raw.ToString() : "(Ô trống)";
                                    sb.AppendLine($"[Cell {cellRef}]: \"{textVal}\"");
                                    addedContext = true;
                                }
                            }
                            catch { }
                        }
                    }

                    // If no specific cell ref in prompt, read current ActiveCell if available
                    if (!addedContext && excelApp.ActiveCell != null)
                    {
                        try
                        {
                            Excel.Range activeCell = excelApp.ActiveCell;
                            object raw = activeCell.Value2;
                            if (raw != null && !string.IsNullOrEmpty(raw.ToString()))
                            {
                                sb.AppendLine("\n--- DỮ LIỆU ĐỌC TỰ ĐỘNG TỪ Ô ĐANG CHỌN ---");
                                sb.AppendLine($"[Cell {activeCell.Address[false, false]}]: \"{raw}\"");
                            }
                        }
                        catch { }
                    }
                }

                return sb.ToString();
            }
            catch
            {
                return userPrompt;
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
