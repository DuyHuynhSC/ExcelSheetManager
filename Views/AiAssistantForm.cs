using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
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
            this.Text = "Local AI Assistant for Excel";
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
            _cmbMode.Items.Add("Generate Excel Formula");
            _cmbMode.Items.Add("Translate Cell / Range (Dịch thuật)");
            _cmbMode.Items.Add("Explain Sheet / Formula");
            _cmbMode.Items.Add("General Excel Assistant");
            _cmbMode.SelectedIndex = 1; // Default to Translate mode

            _btnSettings = new Button
            {
                Text = "Settings",
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
                Text = "Hãy dịch nội dung tại cell đang chọn từ tiếng Nhật sang tiếng Việt"
            };

            _btnSend = new Button
            {
                Text = "Send",
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

            // Bottom Panel (TableLayoutPanel prevents _lblStatus from overlapping buttons)
            TableLayoutPanel pnlBottom = new TableLayoutPanel
            {
                Dock = DockStyle.Bottom,
                Height = 46,
                ColumnCount = 3,
                RowCount = 1,
                Padding = new Padding(8, 4, 8, 4),
                Margin = new Padding(0)
            };
            pnlBottom.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f)); // Status label gets remaining space
            pnlBottom.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 150f)); // Insert into Cell button
            pnlBottom.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 115f)); // Copy Text button
            pnlBottom.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));

            _lblStatus = new Label
            {
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft,
                AutoEllipsis = true,
                Text = "Specify cell (e.g. B1) in prompt for auto-reading Excel data.",
                ForeColor = Color.FromArgb(148, 163, 184)
            };

            _btnInsertCell = new Button
            {
                Dock = DockStyle.Fill,
                Margin = new Padding(2),
                Text = "Insert into Cell",
                BackColor = Color.FromArgb(16, 185, 129),
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 9f, FontStyle.Bold),
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            _btnInsertCell.FlatAppearance.BorderSize = 0;
            _btnInsertCell.Click += (s, e) => InsertIntoActiveCell();

            _btnCopy = new Button
            {
                Dock = DockStyle.Fill,
                Margin = new Padding(2),
                Text = "Copy Text",
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            _btnCopy.Click += (s, e) => CopyResponse();

            pnlBottom.Controls.Add(_lblStatus, 0, 0);
            pnlBottom.Controls.Add(_btnInsertCell, 1, 0);
            pnlBottom.Controls.Add(_btnCopy, 2, 0);

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
            var (enrichedPrompt, logMessage) = EnrichPromptWithExcelCellData(prompt);
            _lblStatus.Text = logMessage;

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
                _lblStatus.Text = string.IsNullOrEmpty(logMessage) ? "AI Response received successfully." : logMessage;
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

        private (string EnrichedPrompt, string LogMessage) EnrichPromptWithExcelCellData(string userPrompt)
        {
            try
            {
                var excelApp = (Excel.Application)ExcelDnaUtil.Application;
                if (excelApp == null) return (userPrompt, "Excel connection unavailable.");

                // Regex matching cell references like B1, A1:C10 (optional prefix cell/ô)
                var matches = Regex.Matches(
                    userPrompt,
                    @"(?:cell|ô|o)?\s*\b([A-Za-z]{1,3}[1-9][0-9]{0,5}(?::[A-Za-z]{1,3}[1-9][0-9]{0,5})?)\b",
                    RegexOptions.IgnoreCase
                );

                List<string> foundRefs = new List<string>();
                foreach (Match m in matches)
                {
                    if (m.Groups.Count > 1 && !string.IsNullOrEmpty(m.Groups[1].Value))
                    {
                        string cellRef = m.Groups[1].Value.Trim().ToUpper();
                        // Filter out common non-cell short words
                        if (cellRef.Length == 1 || (cellRef.Length <= 3 && !cellRef.Any(char.IsDigit))) continue;

                        if (!foundRefs.Contains(cellRef, StringComparer.OrdinalIgnoreCase))
                        {
                            foundRefs.Add(cellRef);
                        }
                    }
                }

                StringBuilder sb = new StringBuilder();
                sb.AppendLine(userPrompt);

                List<string> readSummary = new List<string>();

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
                                Excel.Range? range = activeWs.Range[cellRef];
                                if (range != null)
                                {
                                    object raw = range.Value2 ?? range.Value;
                                    string textVal = raw != null ? raw.ToString() : "(Ô trống)";
                                    sb.AppendLine($"[Cell {cellRef}]: \"{textVal}\"");
                                    readSummary.Add($"{cellRef}=\"{textVal}\"");
                                    addedContext = true;
                                }
                            }
                            catch (Exception ex)
                            {
                                readSummary.Add($"{cellRef}=(Error: {ex.Message})");
                            }
                        }
                    }

                    // If no specific cell ref in prompt, read current ActiveCell if available
                    if (!addedContext && excelApp.ActiveCell != null)
                    {
                        try
                        {
                            Excel.Range activeCell = excelApp.ActiveCell;
                            object raw = activeCell.Value2 ?? activeCell.Value;
                            if (raw != null && !string.IsNullOrEmpty(raw.ToString()))
                            {
                                string addr = activeCell.Address[false, false];
                                sb.AppendLine("\n--- DỮ LIỆU ĐỌC TỰ ĐỘNG TỪ Ô ĐANG CHỌN ---");
                                sb.AppendLine($"[Cell {addr}]: \"{raw}\"");
                                readSummary.Add($"Active {addr}=\"{raw}\"");
                            }
                        }
                        catch { }
                    }
                }

                string statusMsg = readSummary.Count > 0
                    ? $"Read Excel: {string.Join(", ", readSummary)}"
                    : "No Excel cells referenced or active cell is empty.";

                return (sb.ToString(), statusMsg);
            }
            catch (Exception ex)
            {
                return (userPrompt, $"Excel read error: {ex.Message}");
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
