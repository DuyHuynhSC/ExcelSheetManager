using System;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;
using ExcelSheetManager.Helpers;

namespace ExcelSheetManager.Views
{
    internal class GlossaryForm : Form
    {
        private TextBox _txtGlossary = null!;
        private Button _btnSave = null!;
        private Button _btnImport = null!;
        private Button _btnExport = null!;
        private Button _btnResetSample = null!;
        private Label _lblInstruction = null!;
        private Label _lblStatus = null!;

        public GlossaryForm()
        {
            this.Size = new Size(580, 520);
            this.Text = "Glossary (JP - VN) - Từ Điển Thuật Ngữ Dịch AI";
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.TopMost = true;

            bool isDark = SettingsHelper.GetIsDarkTheme();
            this.BackColor = isDark ? Color.FromArgb(15, 23, 42) : Color.FromArgb(241, 245, 249);
            this.ForeColor = isDark ? Color.FromArgb(248, 250, 252) : Color.FromArgb(15, 23, 42);
            this.Font = new Font("Segoe UI", 9.5f);

            InitializeComponent();
            LoadGlossary();
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();

            bool isDark = SettingsHelper.GetIsDarkTheme();

            _lblInstruction = new Label
            {
                Location = new Point(16, 12),
                Size = new Size(530, 42),
                Text = "Nhập danh sách thuật ngữ Tiếng Nhật = Nghĩa Tiếng Việt (mỗi cặp 1 dòng).\nVí dụ: 売上=Doanh thu",
                Font = new Font("Segoe UI", 9f, FontStyle.Bold),
                ForeColor = Color.FromArgb(59, 130, 246)
            };

            _txtGlossary = new TextBox
            {
                Location = new Point(16, 58),
                Size = new Size(530, 330),
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                Font = new Font("Consolas", 10f),
                BackColor = isDark ? Color.FromArgb(30, 41, 59) : Color.White,
                ForeColor = isDark ? Color.FromArgb(248, 250, 252) : Color.FromArgb(15, 23, 42)
            };

            _lblStatus = new Label
            {
                Location = new Point(16, 396),
                Size = new Size(530, 20),
                Text = "Sẵn sàng. Nhập thủ công hoặc bấm 'Import từ File...' để nạp dữ liệu.",
                Font = new Font("Segoe UI", 8.5f),
                ForeColor = Color.FromArgb(148, 163, 184)
            };

            // Buttons
            _btnImport = new Button
            {
                Location = new Point(16, 426),
                Size = new Size(125, 32),
                Text = "Import từ File...",
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            _btnImport.Click += (s, e) => ImportFromFile();

            _btnExport = new Button
            {
                Location = new Point(148, 426),
                Size = new Size(125, 32),
                Text = "Export ra File...",
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            _btnExport.Click += (s, e) => ExportToFile();

            _btnResetSample = new Button
            {
                Location = new Point(280, 426),
                Size = new Size(135, 32),
                Text = "Nạp Thuật Ngữ Mẫu",
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            _btnResetSample.Click += (s, e) => LoadSampleTerms();

            _btnSave = new Button
            {
                Location = new Point(422, 426),
                Size = new Size(124, 32),
                Text = "Save & Close",
                BackColor = Color.FromArgb(16, 185, 129),
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 9f, FontStyle.Bold),
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            _btnSave.FlatAppearance.BorderSize = 0;
            _btnSave.Click += (s, e) => SaveGlossary();

            this.Controls.Add(_lblInstruction);
            this.Controls.Add(_txtGlossary);
            this.Controls.Add(_lblStatus);
            this.Controls.Add(_btnImport);
            this.Controls.Add(_btnExport);
            this.Controls.Add(_btnResetSample);
            this.Controls.Add(_btnSave);

            this.ResumeLayout(false);
        }

        private void LoadGlossary()
        {
            _txtGlossary.Text = SettingsHelper.GetGlossaryText();
        }

        private void LoadSampleTerms()
        {
            string sample = "売上=Doanh thu\r\n利益=Lợi nhuận\r\n勘定科目=Tài khoản kế toán\r\n残高=Số dư\r\n振込=Chuyển khoản\r\n税込=Đã bao gồm thuế\r\n税抜=Chưa bao gồm thuế\r\n請求書=Hóa đơn\r\n納品書=Biên bản giao hàng\r\n発注書=Đơn đặt hàng";
            
            if (string.IsNullOrEmpty(_txtGlossary.Text.Trim()))
            {
                _txtGlossary.Text = sample;
            }
            else
            {
                _txtGlossary.Text += "\r\n" + sample;
            }

            _lblStatus.Text = "Đã nạp thêm 10 thuật ngữ mẫu kế toán/doanh nghiệp!";
            _lblStatus.ForeColor = Color.FromArgb(16, 185, 129);
        }

        private void ImportFromFile()
        {
            try
            {
                using (OpenFileDialog ofd = new OpenFileDialog())
                {
                    ofd.Filter = "Text or CSV Files (*.txt;*.csv)|*.txt;*.csv|All Files (*.*)|*.*";
                    ofd.Title = "Chọn file thuật ngữ Glossary (TXT hoặc CSV)";
                    if (ofd.ShowDialog() == DialogResult.OK)
                    {
                        string[] lines = File.ReadAllLines(ofd.FileName, Encoding.UTF8);
                        StringBuilder sbAppended = new StringBuilder();
                        int count = 0;

                        foreach (string line in lines)
                        {
                            string trimmed = line.Trim();
                            if (string.IsNullOrEmpty(trimmed) || trimmed.StartsWith("#")) continue;

                            // Support both '=' and ',' or Tab delimiters
                            string formatted = trimmed;
                            if (!formatted.Contains("=") && formatted.Contains(","))
                            {
                                string[] parts = formatted.Split(new[] { ',' }, 2);
                                formatted = $"{parts[0].Trim()}={parts[1].Trim()}";
                            }
                            else if (!formatted.Contains("=") && formatted.Contains("\t"))
                            {
                                string[] parts = formatted.Split(new[] { '\t' }, 2);
                                formatted = $"{parts[0].Trim()}={parts[1].Trim()}";
                            }

                            sbAppended.AppendLine(formatted);
                            count++;
                        }

                        if (count > 0)
                        {
                            if (string.IsNullOrEmpty(_txtGlossary.Text.Trim()))
                            {
                                _txtGlossary.Text = sbAppended.ToString();
                            }
                            else
                            {
                                _txtGlossary.Text += "\r\n" + sbAppended.ToString();
                            }

                            _lblStatus.Text = $"Đã nạp thành công {count} thuật ngữ từ file '{Path.GetFileName(ofd.FileName)}'!";
                            _lblStatus.ForeColor = Color.FromArgb(16, 185, 129);
                        }
                        else
                        {
                            _lblStatus.Text = "Không tìm thấy thuật ngữ hợp lệ trong file.";
                            _lblStatus.ForeColor = Color.FromArgb(239, 68, 68);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi đọc file: {ex.Message}", "Import Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ExportToFile()
        {
            try
            {
                using (SaveFileDialog sfd = new SaveFileDialog())
                {
                    sfd.Filter = "Text File (*.txt)|*.txt|CSV File (*.csv)|*.csv";
                    sfd.FileName = "Glossary_JP_VN.txt";
                    sfd.Title = "Xuất danh sách thuật ngữ Glossary ra file";
                    if (sfd.ShowDialog() == DialogResult.OK)
                    {
                        File.WriteAllText(sfd.FileName, _txtGlossary.Text.Trim(), Encoding.UTF8);
                        _lblStatus.Text = $"Đã xuất file Glossary thành công: {Path.GetFileName(sfd.FileName)}";
                        _lblStatus.ForeColor = Color.FromArgb(16, 185, 129);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi xuất file: {ex.Message}", "Export Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void SaveGlossary()
        {
            try
            {
                SettingsHelper.SetGlossaryText(_txtGlossary.Text.Trim());
                _lblStatus.Text = "Đã lưu danh sách Glossary (JP-VN) thành công!";
                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi lưu Glossary: {ex.Message}", "Save Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
