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
        private RichTextBox _rtbResponse = null!;
        private ComboBox _cmbMode = null!;
        private Button _btnSend = null!;
        private Button _btnCopy = null!;
        private Button _btnInsertCell = null!;
        private Button _btnGlossary = null!;
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

            // Top Control Panel
            Panel pnlTop = new Panel
            {
                Dock = DockStyle.Top,
                Height = 130,
                Padding = new Padding(12, 10, 12, 6)
            };

            Label lblMode = new Label { Text = "Mode:", Location = new Point(12, 10), AutoSize = true, Font = new Font("Segoe UI", 9f, FontStyle.Bold) };
            _cmbMode = new ComboBox
            {
                Location = new Point(65, 6),
                Width = 230,
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            _txtPrompt = new TextBox
            {
                Location = new Point(12, 40),
                Width = 420,
                Height = 80,
                Multiline = true,
                ScrollBars = ScrollBars.Vertical
            };

            _cmbMode.Items.Add("Generate Excel Formula");
            _cmbMode.Items.Add("Translate Cell / Range (Dịch thuật)");
            _cmbMode.Items.Add("Explain Sheet / Formula");
            _cmbMode.Items.Add("General Excel Assistant");
            _cmbMode.SelectedIndexChanged += (s, e) => UpdatePromptTemplateForMode();
            _cmbMode.SelectedIndex = 1; // Default to Translate mode

            _btnGlossary = new Button
            {
                Text = "Glossary (JP-VN)",
                Location = new Point(305, 6),
                Width = 125,
                Height = 26,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            _btnGlossary.Click += (s, e) => OpenGlossary();

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

            _btnSend = new Button
            {
                Text = "Send",
                Location = new Point(440, 40),
                Width = 110,
                Height = 80,
                Font = new Font("Segoe UI", 10.5f, FontStyle.Bold),
                ForeColor = Color.White,
                BackColor = Color.FromArgb(59, 130, 246),
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            _btnSend.FlatAppearance.BorderSize = 0;
            _btnSend.Click += async (s, e) => await SendPromptAsync();

            pnlTop.Controls.Add(lblMode);
            pnlTop.Controls.Add(_cmbMode);
            pnlTop.Controls.Add(_btnGlossary);
            pnlTop.Controls.Add(_btnSettings);
            pnlTop.Controls.Add(_txtPrompt);
            pnlTop.Controls.Add(_btnSend);

            // Bottom Action Bar
            TableLayoutPanel pnlBottom = new TableLayoutPanel
            {
                Dock = DockStyle.Bottom,
                Height = 44,
                Padding = new Padding(12, 6, 12, 8),
                ColumnCount = 3,
                RowCount = 1
            };
            pnlBottom.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
            pnlBottom.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 150f));
            pnlBottom.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 115f));

            _lblStatus = new Label
            {
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new Font("Segoe UI", 8.5f),
                Text = "Ready",
                AutoEllipsis = true
            };

            _btnInsertCell = new Button
            {
                Dock = DockStyle.Fill,
                Height = 30,
                Text = "Insert into Cell",
                Font = new Font("Segoe UI", 9f, FontStyle.Bold),
                ForeColor = Color.White,
                BackColor = Color.FromArgb(16, 185, 129),
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            _btnInsertCell.FlatAppearance.BorderSize = 0;
            _btnInsertCell.Click += (s, e) => InsertResponseToCell();

            _btnCopy = new Button
            {
                Dock = DockStyle.Fill,
                Height = 30,
                Text = "Copy Text",
                Font = new Font("Segoe UI", 9f, FontStyle.Bold),
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            _btnCopy.Click += (s, e) => CopyResponse();

            pnlBottom.Controls.Add(_lblStatus, 0, 0);
            pnlBottom.Controls.Add(_btnInsertCell, 1, 0);
            pnlBottom.Controls.Add(_btnCopy, 2, 0);

            // Middle Response Panel (RichTextBox for Perfect Paragraph & Line Break Rendering!)
            _rtbResponse = new RichTextBox
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                BorderStyle = BorderStyle.None,
                ScrollBars = RichTextBoxScrollBars.Vertical,
                Font = new Font("Segoe UI", 9.5f),
                BackColor = _isDarkTheme ? Color.FromArgb(30, 41, 59) : Color.White,
                ForeColor = _isDarkTheme ? Color.FromArgb(248, 250, 252) : Color.FromArgb(15, 23, 42)
            };

            this.Controls.Add(_rtbResponse);
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

        private void OpenGlossary()
        {
            using (GlossaryForm form = new GlossaryForm())
            {
                form.TopMost = true;
                form.ShowDialog(this);
            }
        }

        private void UpdatePromptTemplateForMode()
        {
            if (_txtPrompt == null) return;
            _txtPrompt.Text = _cmbMode.SelectedIndex switch
            {
                0 => "Viết công thức tính tổng cột B từ B2 đến B50 nếu cột A từ A2 đến A50 bằng 'Đã thanh toán'",
                1 => "Hãy dịch nội dung tại cell đang chọn từ tiếng Nhật sang tiếng Việt",
                2 => "Hãy giải thích công thức hoặc ý nghĩa dữ liệu tại ô đang chọn",
                _ => "Cách lọc danh sách các hàng trùng lặp và tính tổng tự động trong Excel là gì?"
            };
        }

        private async Task SendPromptAsync()
        {
            string prompt = _txtPrompt.Text.Trim();
            if (string.IsNullOrEmpty(prompt)) return;

            _btnSend.Enabled = false;
            _lblStatus.Text = "Reading Excel cells & sending to Local AI...";
            _lblStatus.ForeColor = Color.FromArgb(59, 130, 246);
            _rtbResponse.Text = "Reading Excel cells and generating AI response...";

            // Auto-extract cell content referenced in user prompt
            var (enrichedPrompt, logMessage) = EnrichPromptWithExcelCellData(prompt);
            _lblStatus.Text = logMessage;

            string systemPrompt;
            if (_cmbMode.SelectedIndex == 1) // Translate Cell / Range Mode
            {
                StringBuilder sbSystem = new StringBuilder();
                sbSystem.AppendLine("You are an expert Japanese-to-Vietnamese translator integrated inside Microsoft Excel.");
                sbSystem.AppendLine("Read the provided cell content carefully and translate it accurately into natural Vietnamese as requested.");

                string glossaryText = SettingsHelper.GetGlossaryText();
                if (!string.IsNullOrEmpty(glossaryText))
                {
                    sbSystem.AppendLine("\nGLOSSARY / TERMINOLOGY DICTIONARY (STRICTLY ENFORCE THESE TERM TRANSLATIONS):");
                    string[] lines = glossaryText.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                    foreach (string line in lines)
                    {
                        string trimmed = line.Trim();
                        if (trimmed.Contains("="))
                        {
                            string[] parts = trimmed.Split(new[] { '=' }, 2);
                            if (parts.Length == 2 && !string.IsNullOrEmpty(parts[0]) && !string.IsNullOrEmpty(parts[1]))
                            {
                                sbSystem.AppendLine($"- Translate \"{parts[0].Trim()}\" as \"{parts[1].Trim()}\"");
                            }
                        }
                    }
                }

                sbSystem.AppendLine("Output ONLY the translated text without extra conversational fluff.");
                systemPrompt = sbSystem.ToString();
            }
            else
            {
                systemPrompt = _cmbMode.SelectedIndex switch
                {
                    0 => "You are an expert Excel formula generator. Output ONLY the exact Excel formula starting with =. Do not include markdown code block formatting or explanation unless asked.",
                    2 => "You are an expert Excel data analyst. Explain the selected Excel formula in clear, concise Vietnamese.\n\nCRITICAL OUTPUT FORMAT REQUIREMENTS:\n- DO NOT output giant walls of text.\n- Use short paragraphs and bullet points.\n- Format headers with 📌.\n\nUse this structure:\n📌 Ý NGHĨA CÔNG THỨC:\n[1-2 câu giải thích ngắn gọn]\n\n📌 CHI TIẾT CÁC HÀM:\n- [Hàm 1]: [Giải thích ngắn gọn 1 dòng]\n- [Hàm 2]: [Giải thích ngắn gọn 1 dòng]\n\n📌 TÓM TẮT KẾT QUẢ:\n- [Kết quả và ghi chú]",
                    _ => "You are an intelligent, helpful AI assistant integrated inside Microsoft Excel. Provide concise, structured answers in Vietnamese with bullet points."
                };
            }

            try
            {
                string result = await AiService.GetCompletionAsync(enrichedPrompt, systemPrompt);
                RenderFormattedResponse(result);
                _lblStatus.Text = string.IsNullOrEmpty(logMessage) ? "AI Response received successfully." : logMessage;
                _lblStatus.ForeColor = Color.FromArgb(16, 185, 129);
            }
            catch (Exception ex)
            {
                _rtbResponse.Text = $"Error: {ex.Message}";
                _lblStatus.Text = "Error communicating with AI.";
                _lblStatus.ForeColor = Color.FromArgb(239, 68, 68);
            }
            finally
            {
                _btnSend.Enabled = true;
            }
        }

        private void RenderFormattedResponse(string rawText)
        {
            _rtbResponse.Clear();
            if (string.IsNullOrEmpty(rawText)) return;

            _rtbResponse.SuspendLayout();
            try
            {
                // Normalize line endings
                string normalized = rawText.Replace("\r\n", "\n").Replace("\r", "\n");

                // Remove markdown bold and backticks
                normalized = normalized.Replace("**", "").Replace("`", "");

                string[] lines = normalized.Split('\n');
                bool isFirstLine = true;

                foreach (string rawLine in lines)
                {
                    string line = rawLine.TrimEnd();
                    if (string.IsNullOrWhiteSpace(line))
                    {
                        _rtbResponse.AppendText("\r\n\r\n");
                        continue;
                    }

                    if (!isFirstLine)
                    {
                        _rtbResponse.AppendText("\r\n");
                    }
                    isFirstLine = false;

                    bool isHeader = Regex.IsMatch(line, @"^#{1,6}\s") || line.StartsWith("📌") || line.StartsWith("🔹") || Regex.IsMatch(line, @"^\d+\.\s");
                    string cleanLine = Regex.Replace(line, @"^#{1,6}\s*", "📌 ");
                    cleanLine = Regex.Replace(cleanLine, @"^\s*---+\s*$", "──────────────────────────────────────");
                    cleanLine = Regex.Replace(cleanLine, @"^\s*[\*\-]\s+", "  • ");

                    int start = _rtbResponse.TextLength;
                    _rtbResponse.AppendText(cleanLine);
                    int len = cleanLine.Length;

                    if (isHeader)
                    {
                        _rtbResponse.Select(start, len);
                        _rtbResponse.SelectionFont = new Font("Segoe UI", 10f, FontStyle.Bold);
                        _rtbResponse.SelectionColor = _isDarkTheme ? Color.FromArgb(56, 189, 248) : Color.FromArgb(2, 132, 199);
                    }
                    else
                    {
                        _rtbResponse.Select(start, len);
                        _rtbResponse.SelectionFont = new Font("Segoe UI", 9.5f, FontStyle.Regular);
                        _rtbResponse.SelectionColor = _isDarkTheme ? Color.FromArgb(248, 250, 252) : Color.FromArgb(15, 23, 42);
                    }
                }

                _rtbResponse.SelectionStart = 0;
                _rtbResponse.SelectionLength = 0;
                _rtbResponse.ScrollToCaret();
            }
            finally
            {
                _rtbResponse.ResumeLayout();
            }
        }

        private static (string Formula, string Value) ExtractFormulaAndValueFromRange(Excel.Range? range)
        {
            if (range == null) return (string.Empty, string.Empty);
            try
            {
                Excel.Range targetCell = range;
                try
                {
                    if (range.Cells != null && range.Cells.Count > 0)
                    {
                        targetCell = range.Cells[1, 1];
                    }
                }
                catch { }

                // 1. Extract Value
                string textVal = string.Empty;
                try
                {
                    object raw = targetCell.Value2 ?? targetCell.Value;
                    if (raw != null)
                    {
                        if (raw is Array arr)
                        {
                            foreach (object item in arr)
                            {
                                if (item != null && !string.IsNullOrEmpty(item.ToString()))
                                {
                                    textVal = item.ToString();
                                    break;
                                }
                            }
                        }
                        else
                        {
                            textVal = raw.ToString();
                        }
                    }
                }
                catch { }

                if (string.IsNullOrEmpty(textVal)) textVal = "(Ô trống)";

                // 2. Extract Formula
                string formulaVal = string.Empty;

                // Try targetCell.Formula2
                try
                {
                    object f2 = targetCell.Formula2;
                    if (f2 != null)
                    {
                        if (f2 is Array fArr)
                        {
                            foreach (object item in fArr)
                            {
                                if (item != null && item.ToString().Trim().StartsWith("="))
                                {
                                    formulaVal = item.ToString().Trim();
                                    break;
                                }
                            }
                        }
                        else
                        {
                            string s = f2.ToString().Trim();
                            if (s.StartsWith("=")) formulaVal = s;
                        }
                    }
                }
                catch { }

                // Try targetCell.Formula
                if (string.IsNullOrEmpty(formulaVal))
                {
                    try
                    {
                        object f1 = targetCell.Formula;
                        if (f1 != null)
                        {
                            if (f1 is Array fArr)
                            {
                                foreach (object item in fArr)
                                {
                                    if (item != null && item.ToString().Trim().StartsWith("="))
                                    {
                                        formulaVal = item.ToString().Trim();
                                        break;
                                    }
                                }
                            }
                            else
                            {
                                string s = f1.ToString().Trim();
                                if (s.StartsWith("=")) formulaVal = s;
                            }
                        }
                    }
                    catch { }
                }

                // Try targetCell.FormulaLocal
                if (string.IsNullOrEmpty(formulaVal))
                {
                    try
                    {
                        object fl = targetCell.FormulaLocal;
                        if (fl != null)
                        {
                            if (fl is Array fArr)
                            {
                                foreach (object item in fArr)
                                {
                                    if (item != null && item.ToString().Trim().StartsWith("="))
                                    {
                                        formulaVal = item.ToString().Trim();
                                        break;
                                    }
                                }
                            }
                            else
                            {
                                string s = fl.ToString().Trim();
                                if (s.StartsWith("=")) formulaVal = s;
                            }
                        }
                    }
                    catch { }
                }

                return (formulaVal, textVal);
            }
            catch { }
            return (string.Empty, string.Empty);
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
                        sb.AppendLine("\n--- DỮ LIỆU VÀ CÔNG THỨC ĐỌC TỰ ĐỘNG TỪ EXCEL ---");
                        foreach (string cellRef in foundRefs)
                        {
                            try
                            {
                                Excel.Range? range = activeWs.Range[cellRef];
                                if (range != null)
                                {
                                    var (formulaVal, textVal) = ExtractFormulaAndValueFromRange(range);

                                    if (!string.IsNullOrEmpty(formulaVal))
                                    {
                                        sb.AppendLine($"[Cell {cellRef}]:");
                                        sb.AppendLine($"  - FORMULA: {formulaVal}");
                                        sb.AppendLine($"  - EVALUATED VALUE: {textVal}");
                                        readSummary.Add($"{cellRef} (Formula: {formulaVal})");
                                    }
                                    else
                                    {
                                        sb.AppendLine($"[Cell {cellRef}]: Value = \"{textVal}\"");
                                        readSummary.Add($"{cellRef}=\"{textVal}\"");
                                    }
                                    addedContext = true;
                                }
                            }
                            catch (Exception ex)
                            {
                                readSummary.Add($"{cellRef}=(Error: {ex.Message})");
                            }
                        }
                    }

                    // If no specific cell ref in prompt, read current ActiveCell or Selection if available
                    if (!addedContext)
                    {
                        Excel.Range? activeCell = null;
                        try { activeCell = excelApp.ActiveCell; } catch { }
                        if (activeCell == null)
                        {
                            try { activeCell = excelApp.Selection as Excel.Range; } catch { }
                        }

                        if (activeCell != null)
                        {
                            try
                            {
                                var (formulaVal, textVal) = ExtractFormulaAndValueFromRange(activeCell);
                                string addr = activeCell.Address[false, false];

                                sb.AppendLine("\n--- DỮ LIỆU VÀ CÔNG THỨC ĐỌC TỰ ĐỘNG TỪ Ô ĐANG CHỌN ---");
                                if (!string.IsNullOrEmpty(formulaVal))
                                {
                                    sb.AppendLine($"[Cell {addr}]:");
                                    sb.AppendLine($"  - FORMULA: {formulaVal}");
                                    sb.AppendLine($"  - EVALUATED VALUE: {textVal}");
                                    readSummary.Add($"Active {addr} (Formula: {formulaVal})");
                                }
                                else
                                {
                                    sb.AppendLine($"[Cell {addr}]: Value = \"{textVal}\"");
                                    readSummary.Add($"Active {addr}=\"{textVal}\"");
                                }
                            }
                            catch { }
                        }
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
            if (!string.IsNullOrEmpty(_rtbResponse.Text))
            {
                Clipboard.SetText(_rtbResponse.Text);
                _lblStatus.Text = "Copied AI response to clipboard.";
            }
        }

        private void InsertIntoActiveCell()
        {
            try
            {
                string text = _rtbResponse.Text.Trim();
                if (string.IsNullOrEmpty(text)) return;

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

        private void InsertResponseToCell() => InsertIntoActiveCell();
    }
}
