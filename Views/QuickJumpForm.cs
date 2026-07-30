using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelSheetManager.Helpers;
using ExcelSheetManager.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelSheetManager.Views
{
    internal class QuickJumpForm : Form
    {
        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool BringWindowToTop(IntPtr hWnd);

        private static QuickJumpForm? _currentInstance;

        internal static void ShowForm()
        {
            try
            {
                if (_currentInstance != null && !_currentInstance.IsDisposed)
                {
                    _currentInstance.Activate();
                    _currentInstance.FocusSearchBox();
                    return;
                }

                _currentInstance = new QuickJumpForm();
                _currentInstance.Show();
                _currentInstance.Activate();
                _currentInstance.FocusSearchBox();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Could not open Quick Jump window: {ex.Message}", "Quick Jump Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private TextBox _txtSearch = null!;
        private ListBox _lstResults = null!;
        private Label _lblFooterTip = null!;
        private Panel _pnlHeader = null!;
        private Panel _pnlFooter = null!;

        private readonly List<QuickJumpItem> _allItems = new();
        private bool _isDarkTheme = true;

        private Color _bgColor;
        private Color _cardColor;
        private Color _textColor;
        private Color _subTextColor;
        private Color _selectColor;
        private Color _borderColor;

        public QuickJumpForm()
        {
            this.Size = new Size(540, 380);
            this.FormBorderStyle = FormBorderStyle.None;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.TopMost = true;
            this.ShowInTaskbar = false;
            this.KeyPreview = true;
            this.DoubleBuffered = true;

            _isDarkTheme = SettingsHelper.GetIsDarkTheme(defaultValue: true);

            InitializeComponent();
            ApplyTheme();
            LoadExcelData();
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();

            _pnlHeader = new Panel
            {
                Dock = DockStyle.Top,
                Height = 48,
                Padding = new Padding(12, 10, 12, 8)
            };

            _txtSearch = new TextBox
            {
                Dock = DockStyle.Fill,
                Font = new Font("Segoe UI", 11f),
                BorderStyle = BorderStyle.FixedSingle
            };
            _txtSearch.TextChanged += (s, e) => FilterResults();
            _txtSearch.KeyDown += TxtSearch_KeyDown;

            _pnlHeader.Controls.Add(_txtSearch);

            _pnlFooter = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 28,
                Padding = new Padding(12, 4, 12, 4)
            };

            _lblFooterTip = new Label
            {
                Dock = DockStyle.Fill,
                Font = new Font("Segoe UI", 8.5f),
                Text = "↑ ↓ Select  •  Enter Jump  •  Esc Close",
                TextAlign = ContentAlignment.MiddleLeft
            };
            _pnlFooter.Controls.Add(_lblFooterTip);

            _lstResults = new ListBox
            {
                Dock = DockStyle.Fill,
                BorderStyle = BorderStyle.None,
                DrawMode = DrawMode.OwnerDrawFixed,
                ItemHeight = 40
            };
            _lstResults.DrawItem += DrawResultItem;
            _lstResults.DoubleClick += (s, e) => ExecuteJump();
            _lstResults.KeyDown += LstResults_KeyDown;

            this.Controls.Add(_lstResults);
            this.Controls.Add(_pnlHeader);
            this.Controls.Add(_pnlFooter);

            this.ResumeLayout(false);
        }

        public void FocusSearchBox()
        {
            _txtSearch.Focus();
            _txtSearch.SelectAll();
        }

        private void ApplyTheme()
        {
            _bgColor = _isDarkTheme ? Color.FromArgb(15, 23, 42) : Color.FromArgb(241, 245, 249);
            _cardColor = _isDarkTheme ? Color.FromArgb(30, 41, 59) : Color.FromArgb(255, 255, 255);
            _textColor = _isDarkTheme ? Color.FromArgb(248, 250, 252) : Color.FromArgb(15, 23, 42);
            _subTextColor = _isDarkTheme ? Color.FromArgb(148, 163, 184) : Color.FromArgb(100, 116, 139);
            _selectColor = Color.FromArgb(2, 132, 199);
            _borderColor = _isDarkTheme ? Color.FromArgb(51, 65, 85) : Color.FromArgb(203, 213, 225);

            this.BackColor = _bgColor;
            _pnlHeader.BackColor = _cardColor;
            _pnlFooter.BackColor = _cardColor;

            _txtSearch.BackColor = _bgColor;
            _txtSearch.ForeColor = _textColor;

            _lstResults.BackColor = _bgColor;
            _lstResults.ForeColor = _textColor;

            _lblFooterTip.ForeColor = _subTextColor;
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            // Draw 1px border around borderless form
            using (var pen = new Pen(_borderColor, 2))
            {
                e.Graphics.DrawRectangle(pen, 0, 0, this.Width - 1, this.Height - 1);
            }
        }

        private void LoadExcelData()
        {
            _allItems.Clear();
            try
            {
                var excelApp = (Excel.Application)ExcelDnaUtil.Application;
                if (excelApp == null) return;

                foreach (Excel.Workbook wb in excelApp.Workbooks)
                {
                    string wbName = wb.Name;

                    // Add Workbook item
                    _allItems.Add(new QuickJumpItem
                    {
                        Title = wbName,
                        Subtitle = wb.FullName,
                        ItemType = "Workbook",
                        TargetRef = wb,
                        ParentWorkbookName = wbName
                    });

                    // Add Sheets of this workbook
                    foreach (object sheetObj in wb.Sheets)
                    {
                        if (sheetObj is Excel.Worksheet ws)
                        {
                            var (colorHex, hasCustom) = ColorHelper.GetSheetTabColorHex(ws);
                            _allItems.Add(new QuickJumpItem
                            {
                                Title = ws.Name,
                                Subtitle = wbName,
                                ItemType = "Sheet",
                                TabColorHex = colorHex,
                                HasCustomTabColor = hasCustom,
                                IsProtected = ws.ProtectContents,
                                IsHidden = ws.Visible != Excel.XlSheetVisibility.xlSheetVisible,
                                TargetRef = ws,
                                ParentWorkbookName = wbName
                            });
                        }
                        else if (sheetObj is Excel.Chart chart)
                        {
                            _allItems.Add(new QuickJumpItem
                            {
                                Title = chart.Name,
                                Subtitle = wbName,
                                ItemType = "Chart",
                                TabColorHex = "#F59E0B",
                                HasCustomTabColor = true,
                                IsProtected = false,
                                IsHidden = chart.Visible != Excel.XlSheetVisibility.xlSheetVisible,
                                TargetRef = chart,
                                ParentWorkbookName = wbName
                            });
                        }
                    }
                }

                FilterResults();
            }
            catch (Exception ex)
            {
                _lblFooterTip.Text = $"Load error: {ex.Message}";
            }
        }

        private void FilterResults()
        {
            _lstResults.BeginUpdate();
            _lstResults.Items.Clear();

            string filter = _txtSearch.Text.Trim();
            var matches = string.IsNullOrEmpty(filter)
                ? _allItems
                : _allItems.Where(i => i.Title.IndexOf(filter, StringComparison.OrdinalIgnoreCase) >= 0 ||
                                       i.Subtitle.IndexOf(filter, StringComparison.OrdinalIgnoreCase) >= 0);

            foreach (var item in matches)
            {
                _lstResults.Items.Add(item);
            }

            if (_lstResults.Items.Count > 0)
            {
                _lstResults.SelectedIndex = 0;
            }

            _lstResults.EndUpdate();
        }

        private void TxtSearch_KeyDown(object? sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Down)
            {
                if (_lstResults.Items.Count > 0)
                {
                    int next = Math.Min(_lstResults.SelectedIndex + 1, _lstResults.Items.Count - 1);
                    _lstResults.SelectedIndex = next;
                }
                e.Handled = true;
            }
            else if (e.KeyCode == Keys.Up)
            {
                if (_lstResults.Items.Count > 0)
                {
                    int prev = Math.Max(_lstResults.SelectedIndex - 1, 0);
                    _lstResults.SelectedIndex = prev;
                }
                e.Handled = true;
            }
            else if (e.KeyCode == Keys.Enter)
            {
                ExecuteJump();
                e.Handled = true;
            }
            else if (e.KeyCode == Keys.Escape)
            {
                this.Close();
                e.Handled = true;
            }
        }

        private void LstResults_KeyDown(object? sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                ExecuteJump();
                e.Handled = true;
            }
            else if (e.KeyCode == Keys.Escape)
            {
                this.Close();
                e.Handled = true;
            }
        }

        private void ExecuteJump()
        {
            if (_lstResults.SelectedIndex < 0 || _lstResults.SelectedIndex >= _lstResults.Items.Count) return;
            if (_lstResults.Items[_lstResults.SelectedIndex] is not QuickJumpItem item) return;

            try
            {
                if (item.ItemType == "Workbook" && item.TargetRef is Excel.Workbook wb)
                {
                    wb.Activate();
                    if (wb.Windows != null && wb.Windows.Count > 0)
                    {
                        Excel.Window win = (Excel.Window)wb.Windows[1];
                        win.Activate();
                        IntPtr hwnd = new IntPtr(win.Hwnd);
                        if (hwnd != IntPtr.Zero)
                        {
                            BringWindowToTop(hwnd);
                            SetForegroundWindow(hwnd);
                        }
                    }
                    AddIn.SyncSelectedWorkbookInAllTaskPanes(wb.Name);
                }
                else if ((item.ItemType == "Sheet" || item.ItemType == "Chart") && item.TargetRef != null)
                {
                    Excel.Workbook? parentWb = null;
                    if (item.TargetRef is Excel.Worksheet ws)
                    {
                        parentWb = (Excel.Workbook)ws.Parent;
                        parentWb.Activate();
                        if (parentWb.Windows != null && parentWb.Windows.Count > 0)
                        {
                            Excel.Window win = (Excel.Window)parentWb.Windows[1];
                            win.Activate();
                            IntPtr hwnd = new IntPtr(win.Hwnd);
                            if (hwnd != IntPtr.Zero)
                            {
                                BringWindowToTop(hwnd);
                                SetForegroundWindow(hwnd);
                            }
                        }
                        if (ws.Visible != Excel.XlSheetVisibility.xlSheetVisible)
                        {
                            ws.Visible = Excel.XlSheetVisibility.xlSheetVisible;
                        }
                        ws.Activate();
                    }
                    else if (item.TargetRef is Excel.Chart chart)
                    {
                        parentWb = (Excel.Workbook)chart.Parent;
                        parentWb.Activate();
                        if (parentWb.Windows != null && parentWb.Windows.Count > 0)
                        {
                            Excel.Window win = (Excel.Window)parentWb.Windows[1];
                            win.Activate();
                            IntPtr hwnd = new IntPtr(win.Hwnd);
                            if (hwnd != IntPtr.Zero)
                            {
                                BringWindowToTop(hwnd);
                                SetForegroundWindow(hwnd);
                            }
                        }
                        chart.Activate();
                    }

                    if (parentWb != null)
                    {
                        AddIn.SyncSelectedWorkbookInAllTaskPanes(parentWb.Name);
                    }
                }

                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Could not jump to target: {ex.Message}", "Jump Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void DrawResultItem(object? sender, DrawItemEventArgs e)
        {
            if (e.Index < 0 || e.Index >= _lstResults.Items.Count) return;
            var item = (QuickJumpItem)_lstResults.Items[e.Index];

            bool isSelected = (e.Index == _lstResults.SelectedIndex);
            Color itemBg = isSelected ? _selectColor : _cardColor;
            Color titleColor = isSelected ? Color.White : _textColor;
            Color subColor = isSelected ? Color.FromArgb(224, 242, 254) : _subTextColor;

            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;

            // Card Background
            using (var brush = new SolidBrush(itemBg))
            {
                e.Graphics.FillRectangle(brush, new Rectangle(e.Bounds.X + 2, e.Bounds.Y + 2, e.Bounds.Width - 4, e.Bounds.Height - 4));
            }

            // Icon Prefix
            string iconStr = item.ItemType == "Workbook" ? "📄" : "📋";
            TextRenderer.DrawText(e.Graphics, iconStr, new Font("Segoe UI", 11f), new Point(e.Bounds.X + 8, e.Bounds.Y + 8), titleColor);

            // Title (Name)
            using (var fontTitle = new Font("Segoe UI", 9.5f, FontStyle.Bold))
            {
                Rectangle titleRect = new Rectangle(e.Bounds.X + 34, e.Bounds.Y + 4, e.Bounds.Width - 110, 18);
                TextRenderer.DrawText(e.Graphics, item.Title, fontTitle, titleRect, titleColor, TextFormatFlags.Left | TextFormatFlags.EndEllipsis);
            }

            // Subtitle (Workbook / Path)
            using (var fontSub = new Font("Segoe UI", 8.5f))
            {
                Rectangle subRect = new Rectangle(e.Bounds.X + 34, e.Bounds.Y + 22, e.Bounds.Width - 110, 16);
                TextRenderer.DrawText(e.Graphics, item.Subtitle, fontSub, subRect, subColor, TextFormatFlags.Left | TextFormatFlags.EndEllipsis);
            }

            // Lock Badge if Protected
            if (item.IsProtected)
            {
                using (var fontBadge = new Font("Segoe UI", 7.5f, FontStyle.Bold))
                using (var badgeBg = new SolidBrush(Color.FromArgb(245, 158, 11))) // Amber 500
                {
                    Rectangle badgeRect = new Rectangle(e.Bounds.Right - 68, e.Bounds.Y + 10, 28, 18);
                    e.Graphics.FillRectangle(badgeBg, badgeRect);
                    TextRenderer.DrawText(e.Graphics, "🔒", fontBadge, new Point(badgeRect.X + 6, badgeRect.Y + 2), Color.White);
                }
            }

            // Hidden Badge if Hidden
            if (item.IsHidden)
            {
                using (var fontBadge = new Font("Segoe UI", 7.5f, FontStyle.Bold))
                using (var badgeBg = new SolidBrush(Color.FromArgb(239, 68, 68))) // Red 500
                {
                    Rectangle badgeRect = new Rectangle(e.Bounds.Right - 38, e.Bounds.Y + 10, 32, 18);
                    e.Graphics.FillRectangle(badgeBg, badgeRect);
                    TextRenderer.DrawText(e.Graphics, "Hide", fontBadge, new Point(badgeRect.X + 4, badgeRect.Y + 2), Color.White);
                }
            }
        }
    }
}
