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
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    [ProgId("ExcelSheetManager.TaskPaneControl")]
    [Guid("A5F19D34-9B05-4B82-94C3-7E4A8D9183C2")]
    public class TaskPaneControl : UserControl
    {
        [DllImport("user32.dll")]
        private static extern IntPtr SendMessage(IntPtr hWnd, int msg, bool wParam, int lParam);
        private const int WM_SETREDRAW = 0x000B;

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);
        private const int SW_RESTORE = 9;

        public string BoundWorkbookName { get; set; } = string.Empty;

        // UI Colors
        private readonly Color _bgColor = Color.FromArgb(15, 23, 42);      // Slate 900 #0F172A
        private readonly Color _cardColor = Color.FromArgb(30, 41, 59);    // Slate 800 #1E293B
        private readonly Color _hoverColor = Color.FromArgb(51, 65, 85);   // Slate 700 #334155
        private readonly Color _selectColor = Color.FromArgb(2, 132, 199);  // Sky 600 #0284C7
        private readonly Color _textColor = Color.FromArgb(248, 250, 252);  // Slate 50 #F8FAFC
        private readonly Color _subTextColor = Color.FromArgb(148, 163, 184);// Slate 400 #94A3B8

        // Controls
        private Panel _pnlHeader = null!;
        private Label _lblTitle = null!;
        private Label _lblSubTitle = null!;
        private Button _btnRefresh = null!;

        private SplitContainer _splitContainer = null!;

        private Panel _pnlVung1Header = null!;
        private Label _lblVung1Title = null!;
        private TextBox _txtFilterFile = null!;
        private ListBox _lstWorkbooks = null!;

        private Panel _pnlVung2Header = null!;
        private Label _lblVung2Title = null!;
        private TextBox _txtFilterSheet = null!;
        private ListBox _lstSheets = null!;

        private Label _lblStatus = null!;

        // Data Storage
        private readonly List<WorkbookItem> _allWorkbooks = new();
        private readonly List<SheetItem> _allSheets = new();
        private WorkbookItem? _selectedWorkbook;
        private bool _isUpdatingUi = false;

        public TaskPaneControl()
        {
            this.Size = new Size(340, 600);
            this.MinimumSize = new Size(200, 300);
            this.BackColor = _bgColor;
            this.ForeColor = _textColor;
            this.Font = new Font("Segoe UI", 9.5f, FontStyle.Regular);

            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();

            // 1. HEADER PANEL
            _pnlHeader = new Panel
            {
                Dock = DockStyle.Top,
                Height = 60,
                BackColor = _cardColor,
                Padding = new Padding(10)
            };

            _lblTitle = new Label
            {
                Text = "📊 Excel Sheet Manager",
                Font = new Font("Segoe UI", 11f, FontStyle.Bold),
                ForeColor = _textColor,
                AutoSize = true,
                Location = new Point(10, 8)
            };

            _lblSubTitle = new Label
            {
                Text = "Workbook & Sheet Navigator",
                Font = new Font("Segoe UI", 8.5f, FontStyle.Regular),
                ForeColor = _subTextColor,
                AutoSize = true,
                Location = new Point(12, 32)
            };

            _btnRefresh = new Button
            {
                Text = "🔄 Refresh",
                Font = new Font("Segoe UI", 9f, FontStyle.Bold),
                ForeColor = Color.White,
                BackColor = Color.FromArgb(59, 130, 246),
                FlatStyle = FlatStyle.Flat,
                Size = new Size(85, 32),
                Location = new Point(240, 14),
                Anchor = AnchorStyles.Top | AnchorStyles.Right,
                Cursor = Cursors.Hand
            };
            _btnRefresh.FlatAppearance.BorderSize = 0;
            _btnRefresh.Click += (s, e) => RefreshData();

            _pnlHeader.Controls.Add(_lblTitle);
            _pnlHeader.Controls.Add(_lblSubTitle);
            _pnlHeader.Controls.Add(_btnRefresh);

            // 2. STATUS LABEL (FOOTER)
            _lblStatus = new Label
            {
                Dock = DockStyle.Bottom,
                Height = 24,
                BackColor = _cardColor,
                ForeColor = _subTextColor,
                Font = new Font("Segoe UI", 8.5f),
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(8, 0, 0, 0),
                Text = "Ready"
            };

            // 3. SPLIT CONTAINER (VÙNG 1 & VÙNG 2)
            _splitContainer = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Horizontal,
                SplitterDistance = 250,
                SplitterWidth = 6,
                BackColor = _bgColor
            };

            // --- VÙNG 1: OPEN FILES ---
            _pnlVung1Header = new Panel
            {
                Dock = DockStyle.Top,
                Height = 55,
                BackColor = _bgColor,
                Padding = new Padding(6, 4, 6, 4)
            };

            _lblVung1Title = new Label
            {
                Text = "📁 OPEN FILES (0)",
                Font = new Font("Segoe UI", 9.5f, FontStyle.Bold),
                ForeColor = Color.FromArgb(56, 189, 248),
                Dock = DockStyle.Top,
                Height = 22
            };

            _txtFilterFile = new TextBox
            {
                Dock = DockStyle.Bottom,
                BackColor = _cardColor,
                ForeColor = _textColor,
                BorderStyle = BorderStyle.FixedSingle,
                Font = new Font("Segoe UI", 9f),
                Text = ""
            };
            _txtFilterFile.TextChanged += (s, e) => FilterWorkbooksList();

            _pnlVung1Header.Controls.Add(_lblVung1Title);
            _pnlVung1Header.Controls.Add(_txtFilterFile);

            _lstWorkbooks = new ListBox
            {
                Dock = DockStyle.Fill,
                BackColor = _bgColor,
                ForeColor = _textColor,
                BorderStyle = BorderStyle.None,
                DrawMode = DrawMode.OwnerDrawFixed,
                ItemHeight = 44
            };
            _lstWorkbooks.DrawItem += DrawWorkbookItem;
            _lstWorkbooks.SelectedIndexChanged += LstWorkbooks_SelectedIndexChanged;

            _splitContainer.Panel1.Controls.Add(_lstWorkbooks);
            _splitContainer.Panel1.Controls.Add(_pnlVung1Header);

            // --- VÙNG 2: SHEETS LIST ---
            _pnlVung2Header = new Panel
            {
                Dock = DockStyle.Top,
                Height = 55,
                BackColor = _bgColor,
                Padding = new Padding(6, 4, 6, 4)
            };

            _lblVung2Title = new Label
            {
                Text = "📋 SHEETS (0)",
                Font = new Font("Segoe UI", 9.5f, FontStyle.Bold),
                ForeColor = Color.FromArgb(129, 140, 248),
                Dock = DockStyle.Top,
                Height = 22
            };

            _txtFilterSheet = new TextBox
            {
                Dock = DockStyle.Bottom,
                BackColor = _cardColor,
                ForeColor = _textColor,
                BorderStyle = BorderStyle.FixedSingle,
                Font = new Font("Segoe UI", 9f),
                Text = ""
            };
            _txtFilterSheet.TextChanged += (s, e) => FilterSheetsList();

            _pnlVung2Header.Controls.Add(_lblVung2Title);
            _pnlVung2Header.Controls.Add(_txtFilterSheet);

            _lstSheets = new ListBox
            {
                Dock = DockStyle.Fill,
                BackColor = _bgColor,
                ForeColor = _textColor,
                BorderStyle = BorderStyle.None,
                DrawMode = DrawMode.OwnerDrawFixed,
                ItemHeight = 40
            };
            _lstSheets.DrawItem += DrawSheetItem;
            _lstSheets.SelectedIndexChanged += LstSheets_SelectedIndexChanged;

            _splitContainer.Panel2.Controls.Add(_lstSheets);
            _splitContainer.Panel2.Controls.Add(_pnlVung2Header);

            // ADD ALL TO CONTROL
            this.Controls.Add(_splitContainer);
            this.Controls.Add(_pnlHeader);
            this.Controls.Add(_lblStatus);

            this.ResumeLayout(false);
        }

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            RefreshData();
        }

        public void RefreshData()
        {
            try
            {
                var excelApp = (Excel.Application)ExcelDnaUtil.Application;
                if (excelApp == null)
                {
                    _lblStatus.Text = "Cannot connect to Excel";
                    return;
                }

                _isUpdatingUi = true;
                _allWorkbooks.Clear();

                string activeName = BoundWorkbookName;
                if (string.IsNullOrEmpty(activeName) && excelApp.ActiveWorkbook != null)
                {
                    try { activeName = excelApp.ActiveWorkbook.Name; } catch { }
                }

                foreach (Excel.Workbook wb in excelApp.Workbooks)
                {
                    try
                    {
                        bool isActive = !string.IsNullOrEmpty(activeName) && wb.Name.Equals(activeName, StringComparison.OrdinalIgnoreCase);
                        _allWorkbooks.Add(new WorkbookItem
                        {
                            Name = wb.Name,
                            FullPath = wb.FullName,
                            WorkbookRef = wb,
                            SheetCount = wb.Sheets.Count,
                            IsActive = isActive
                        });
                    }
                    catch { }
                }

                _lblVung1Title.Text = $"📁 OPEN FILES ({_allWorkbooks.Count})";
                FilterWorkbooksList();

                // Select current bound/active workbook
                WorkbookItem? target = _allWorkbooks.FirstOrDefault(w => w.IsActive)
                                      ?? _allWorkbooks.FirstOrDefault(w => w.Name.Equals(BoundWorkbookName, StringComparison.OrdinalIgnoreCase))
                                      ?? _allWorkbooks.FirstOrDefault();
                if (target != null)
                {
                    SelectWorkbookItemInList(target);
                }

                _lblStatus.Text = $"Updated at {DateTime.Now:HH:mm:ss} ({_allWorkbooks.Count} files)";
                _isUpdatingUi = false;
            }
            catch (Exception ex)
            {
                _isUpdatingUi = false;
                _lblStatus.Text = $"Refresh error: {ex.Message}";
            }
        }

        private void SelectWorkbookItemInList(WorkbookItem item)
        {
            _selectedWorkbook = item;
            int idx = -1;
            for (int i = 0; i < _lstWorkbooks.Items.Count; i++)
            {
                if (_lstWorkbooks.Items[i] is WorkbookItem wbItem && wbItem.Name.Equals(item.Name, StringComparison.OrdinalIgnoreCase))
                {
                    idx = i;
                    break;
                }
            }

            if (idx >= 0 && _lstWorkbooks.SelectedIndex != idx)
            {
                _lstWorkbooks.SelectedIndex = idx;
                _lstWorkbooks.Invalidate();
            }
            LoadSheetsForWorkbook(item);
        }

        private void FilterWorkbooksList()
        {
            if (_lstWorkbooks.IsHandleCreated)
            {
                SendMessage(_lstWorkbooks.Handle, WM_SETREDRAW, false, 0);
            }

            try
            {
                _lstWorkbooks.BeginUpdate();
                _lstWorkbooks.Items.Clear();

                string filter = _txtFilterFile.Text.Trim();
                var matches = string.IsNullOrEmpty(filter)
                    ? _allWorkbooks
                    : _allWorkbooks.Where(w => w.Name.IndexOf(filter, StringComparison.OrdinalIgnoreCase) >= 0 ||
                                               w.FullPath.IndexOf(filter, StringComparison.OrdinalIgnoreCase) >= 0);

                foreach (var item in matches)
                {
                    _lstWorkbooks.Items.Add(item);
                }
            }
            finally
            {
                _lstWorkbooks.EndUpdate();
                if (_lstWorkbooks.IsHandleCreated)
                {
                    SendMessage(_lstWorkbooks.Handle, WM_SETREDRAW, true, 0);
                    _lstWorkbooks.Invalidate();
                }
            }
        }

        private void LoadSheetsForWorkbook(WorkbookItem? wbItem)
        {
            _allSheets.Clear();
            if (_lstSheets.IsHandleCreated)
            {
                SendMessage(_lstSheets.Handle, WM_SETREDRAW, false, 0);
            }

            try
            {
                _lstSheets.BeginUpdate();
                _lstSheets.Items.Clear();

                if (wbItem == null || wbItem.WorkbookRef is not Excel.Workbook wb)
                {
                    _lblVung2Title.Text = "📋 SHEETS (0)";
                    return;
                }

                Excel.Sheets sheets = wb.Sheets;
                object activeSheetObj = wb.ActiveSheet;
                string activeSheetName = (activeSheetObj is Excel.Worksheet wsActive) ? wsActive.Name : string.Empty;

                foreach (object sheetObj in sheets)
                {
                    if (sheetObj is Excel.Worksheet ws)
                    {
                        var (colorHex, hasCustom) = ColorHelper.GetSheetTabColorHex(ws);
                        _allSheets.Add(new SheetItem
                        {
                            Name = ws.Name,
                            ParentWorkbookName = wb.Name,
                            SheetRef = ws,
                            TabColorHex = colorHex,
                            HasCustomTabColor = hasCustom,
                            IsVisible = ws.Visible == Excel.XlSheetVisibility.xlSheetVisible,
                            IsActive = ws.Name.Equals(activeSheetName, StringComparison.OrdinalIgnoreCase),
                            SheetType = "Worksheet"
                        });
                    }
                    else if (sheetObj is Excel.Chart chart)
                    {
                        _allSheets.Add(new SheetItem
                        {
                            Name = chart.Name,
                            ParentWorkbookName = wb.Name,
                            SheetRef = chart,
                            TabColorHex = "#F59E0B",
                            HasCustomTabColor = true,
                            IsVisible = chart.Visible == Excel.XlSheetVisibility.xlSheetVisible,
                            IsActive = false,
                            SheetType = "Chart"
                        });
                    }
                }

                _lblVung2Title.Text = $"📋 SHEETS IN {wb.Name} ({_allSheets.Count})";
                FilterSheetsList();
            }
            catch (Exception ex)
            {
                _lblStatus.Text = $"Load sheets error: {ex.Message}";
            }
            finally
            {
                _lstSheets.EndUpdate();
                if (_lstSheets.IsHandleCreated)
                {
                    SendMessage(_lstSheets.Handle, WM_SETREDRAW, true, 0);
                    _lstSheets.Invalidate();
                }
            }
        }

        private void FilterSheetsList()
        {
            if (_lstSheets.IsHandleCreated)
            {
                SendMessage(_lstSheets.Handle, WM_SETREDRAW, false, 0);
            }

            try
            {
                _lstSheets.BeginUpdate();
                _lstSheets.Items.Clear();

                string filter = _txtFilterSheet.Text.Trim();
                var matches = string.IsNullOrEmpty(filter)
                    ? _allSheets
                    : _allSheets.Where(s => s.Name.IndexOf(filter, StringComparison.OrdinalIgnoreCase) >= 0);

                foreach (var item in matches)
                {
                    _lstSheets.Items.Add(item);
                }
            }
            finally
            {
                _lstSheets.EndUpdate();
                if (_lstSheets.IsHandleCreated)
                {
                    SendMessage(_lstSheets.Handle, WM_SETREDRAW, true, 0);
                    _lstSheets.Invalidate();
                }
            }
        }

        // --- SINGLE-POINT SELECTION EVENT HANDLERS ---
        private void LstWorkbooks_SelectedIndexChanged(object? sender, EventArgs e)
        {
            if (_isUpdatingUi) return;
            if (_lstWorkbooks.SelectedIndex >= 0 && _lstWorkbooks.SelectedIndex < _lstWorkbooks.Items.Count)
            {
                if (_lstWorkbooks.Items[_lstWorkbooks.SelectedIndex] is WorkbookItem item)
                {
                    _isUpdatingUi = true;
                    _selectedWorkbook = item;
                    _lstWorkbooks.Invalidate();

                    LoadSheetsForWorkbook(item);
                    ActivateWorkbookInExcel(item);
                    _isUpdatingUi = false;
                }
            }
        }

        private void LstSheets_SelectedIndexChanged(object? sender, EventArgs e)
        {
            if (_isUpdatingUi) return;
            if (_lstSheets.SelectedIndex >= 0 && _lstSheets.SelectedIndex < _lstSheets.Items.Count)
            {
                if (_lstSheets.Items[_lstSheets.SelectedIndex] is SheetItem item)
                {
                    _isUpdatingUi = true;
                    _lstSheets.Invalidate();
                    ActivateSheetInExcel(item);
                    _isUpdatingUi = false;
                }
            }
        }

        private void ActivateWorkbookInExcel(WorkbookItem item)
        {
            if (item.WorkbookRef is not Excel.Workbook wb) return;
            try
            {
                wb.Activate();
                if (wb.Windows != null && wb.Windows.Count > 0)
                {
                    Excel.Window win = (Excel.Window)wb.Windows[1];
                    win.Activate();
                    IntPtr hwnd = new IntPtr(win.Hwnd);
                    if (hwnd != IntPtr.Zero)
                    {
                        ShowWindowAsync(hwnd, SW_RESTORE);
                        SetForegroundWindow(hwnd);
                    }
                }
                _lblStatus.Text = $"Activated file: {item.Name}";
            }
            catch (Exception ex)
            {
                _lblStatus.Text = $"Could not activate file: {ex.Message}";
            }
        }

        private void ActivateSheetInExcel(SheetItem item)
        {
            try
            {
                Excel.Workbook? parentWb = null;
                if (item.SheetRef is Excel.Worksheet ws)
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
                            ShowWindowAsync(hwnd, SW_RESTORE);
                            SetForegroundWindow(hwnd);
                        }
                    }
                    ws.Activate();
                }
                else if (item.SheetRef is Excel.Chart chart)
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
                            ShowWindowAsync(hwnd, SW_RESTORE);
                            SetForegroundWindow(hwnd);
                        }
                    }
                    chart.Activate();
                }
                _lblStatus.Text = $"Focused sheet: {item.Name}";
            }
            catch (Exception ex)
            {
                _lblStatus.Text = $"Could not focus sheet: {ex.Message}";
            }
        }

        // --- OWNER DRAW LISTBOX ITEMS ---
        private void DrawWorkbookItem(object? sender, DrawItemEventArgs e)
        {
            if (e.Index < 0 || e.Index >= _lstWorkbooks.Items.Count) return;
            var item = (WorkbookItem)_lstWorkbooks.Items[e.Index];

            bool isSelected = (e.Index == _lstWorkbooks.SelectedIndex) ||
                             (_selectedWorkbook != null && item.Name.Equals(_selectedWorkbook.Name, StringComparison.OrdinalIgnoreCase));
            Color itemBg = isSelected ? _selectColor : _cardColor;

            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;

            // Draw Card Background
            using (var brush = new SolidBrush(itemBg))
            {
                e.Graphics.FillRectangle(brush, new Rectangle(e.Bounds.X + 2, e.Bounds.Y + 2, e.Bounds.Width - 4, e.Bounds.Height - 4));
            }

            // Draw Icon
            TextRenderer.DrawText(e.Graphics, "📄", new Font("Segoe UI", 11f), new Point(e.Bounds.X + 8, e.Bounds.Y + 10), _textColor);

            // Draw File Name
            using (var fontName = new Font("Segoe UI", 9.5f, FontStyle.Bold))
            {
                TextRenderer.DrawText(e.Graphics, item.Name, fontName, new Point(e.Bounds.X + 34, e.Bounds.Y + 4), _textColor);
            }

            // Draw Subtitle / Path
            using (var fontSub = new Font("Segoe UI", 8f, FontStyle.Regular))
            {
                TextRenderer.DrawText(e.Graphics, item.DisplaySubtitle, fontSub, new Rectangle(e.Bounds.X + 34, e.Bounds.Y + 24, e.Bounds.Width - 110, 16), _subTextColor, TextFormatFlags.EndEllipsis);
            }

            // Draw Sheet Count Badge
            string badgeText = $"{item.SheetCount} sheets";
            using (var fontBadge = new Font("Segoe UI", 8f, FontStyle.Bold))
            using (var badgeBg = new SolidBrush(Color.FromArgb(51, 65, 85)))
            {
                var size = TextRenderer.MeasureText(badgeText, fontBadge);
                Rectangle badgeRect = new Rectangle(e.Bounds.Right - size.Width - 12, e.Bounds.Y + 12, size.Width + 8, 20);
                e.Graphics.FillRectangle(badgeBg, badgeRect);
                TextRenderer.DrawText(e.Graphics, badgeText, fontBadge, new Point(badgeRect.X + 4, badgeRect.Y + 2), _textColor);
            }
        }

        private void DrawSheetItem(object? sender, DrawItemEventArgs e)
        {
            if (e.Index < 0 || e.Index >= _lstSheets.Items.Count) return;
            var item = (SheetItem)_lstSheets.Items[e.Index];

            bool isSelected = (e.Index == _lstSheets.SelectedIndex);
            Color itemBg = isSelected ? Color.FromArgb(79, 70, 229) : _cardColor; // Indigo 600 if selected

            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;

            // Draw Card Background
            using (var brush = new SolidBrush(itemBg))
            {
                e.Graphics.FillRectangle(brush, new Rectangle(e.Bounds.X + 2, e.Bounds.Y + 2, e.Bounds.Width - 4, e.Bounds.Height - 4));
            }

            // Draw Tab Color Vertical Strip
            Color tabColor = ColorTranslator.FromHtml(item.TabColorHex);
            using (var tabBrush = new SolidBrush(tabColor))
            {
                e.Graphics.FillRectangle(tabBrush, new Rectangle(e.Bounds.X + 6, e.Bounds.Y + 6, 6, e.Bounds.Height - 12));
            }

            // Draw Sheet Name
            using (var fontName = new Font("Segoe UI", 9.5f, FontStyle.Bold))
            {
                TextRenderer.DrawText(e.Graphics, item.Name, fontName, new Point(e.Bounds.X + 20, e.Bounds.Y + 4), _textColor);
            }

            // Draw Sheet Type
            using (var fontType = new Font("Segoe UI", 8f, FontStyle.Regular))
            {
                TextRenderer.DrawText(e.Graphics, item.SheetType, fontType, new Point(e.Bounds.X + 20, e.Bounds.Y + 22), _subTextColor);
            }

            // Draw Hidden Badge if hidden
            if (!item.IsVisible)
            {
                using (var fontBadge = new Font("Segoe UI", 7.5f, FontStyle.Bold))
                using (var badgeBg = new SolidBrush(Color.FromArgb(239, 68, 68))) // Red 500
                {
                    Rectangle badgeRect = new Rectangle(e.Bounds.Right - 55, e.Bounds.Y + 10, 48, 18);
                    e.Graphics.FillRectangle(badgeBg, badgeRect);
                    TextRenderer.DrawText(e.Graphics, "Hidden", fontBadge, new Point(badgeRect.X + 6, badgeRect.Y + 2), Color.White);
                }
            }
        }
    }
}
