using System;
using System.Drawing;
using System.Windows.Forms;
using ExcelSheetManager.Helpers;

namespace ExcelSheetManager.Views
{
    internal class GlossaryForm : Form
    {
        private TextBox _txtGlossary = null!;
        private Button _btnSave = null!;
        private Button _btnResetSample = null!;
        private Button _btnCancel = null!;
        private Label _lblInstruction = null!;

        public GlossaryForm()
        {
            this.Size = new Size(520, 480);
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
                Size = new Size(475, 42),
                Text = "Nhập danh sách thuật ngữ Tiếng Nhật = Nghĩa Tiếng Việt (mỗi cặp 1 dòng).\nVí dụ: 売上=Doanh thu",
                Font = new Font("Segoe UI", 9f, FontStyle.Bold),
                ForeColor = Color.FromArgb(59, 130, 246)
            };

            _txtGlossary = new TextBox
            {
                Location = new Point(16, 58),
                Size = new Size(475, 320),
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                Font = new Font("Consolas", 10f),
                BackColor = isDark ? Color.FromArgb(30, 41, 59) : Color.White,
                ForeColor = isDark ? Color.FromArgb(248, 250, 252) : Color.FromArgb(15, 23, 42)
            };

            _btnResetSample = new Button
            {
                Location = new Point(16, 392),
                Size = new Size(160, 32),
                Text = "Nạp Thuật Ngữ Mẫu",
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            _btnResetSample.Click += (s, e) => LoadSampleTerms();

            _btnSave = new Button
            {
                Location = new Point(275, 392),
                Size = new Size(115, 32),
                Text = "Save",
                BackColor = Color.FromArgb(16, 185, 129),
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 9f, FontStyle.Bold),
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            _btnSave.FlatAppearance.BorderSize = 0;
            _btnSave.Click += (s, e) => SaveGlossary();

            _btnCancel = new Button
            {
                Location = new Point(400, 392),
                Size = new Size(90, 32),
                Text = "Cancel",
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            _btnCancel.Click += (s, e) => this.Close();

            this.Controls.Add(_lblInstruction);
            this.Controls.Add(_txtGlossary);
            this.Controls.Add(_btnResetSample);
            this.Controls.Add(_btnSave);
            this.Controls.Add(_btnCancel);

            this.ResumeLayout(false);
        }

        private void LoadGlossary()
        {
            _txtGlossary.Text = SettingsHelper.GetGlossaryText();
        }

        private void LoadSampleTerms()
        {
            _txtGlossary.Text = "売上=Doanh thu\r\n利益=Lợi nhuận\r\n勘定科目=Tài khoản kế toán\r\n残高=Số dư\r\n振込=Chuyển khoản\r\n税込=Đã bao gồm thuế\r\n税抜=Chưa bao gồm thuế\r\n請求書=Hóa đơn\r\n納品書=Biên bản giao hàng\r\n発注書=Đơn đặt hàng";
        }

        private void SaveGlossary()
        {
            try
            {
                SettingsHelper.SetGlossaryText(_txtGlossary.Text.Trim());
                MessageBox.Show("Đã lưu danh sách Glossary (JP-VN) thành công!", "Glossary Saved", MessageBoxButtons.OK, MessageBoxIcon.Information);
                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi lưu Glossary: {ex.Message}", "Save Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
