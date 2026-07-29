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
        private static extern bool BringWindowToTop(IntPtr hWnd);

        public string BoundWorkbookName { get; set; } = string.Empty;

        // UI Colors
        private readonly Color _bgColor = Color.FromArgb(15, 23, 42);      // Slate 900 #0F172A
        private readonly Color _cardColor = Color.FromArgb(30, 41, 59);    // Slate 800 #1E293B
        private readonly Color _selectColor = Color.FromArgb(2, 132, 199);  // Sky 600 #0284C7
        private readonly Color _textColor = Color.FromArgb(248, 250, 252);  // Slate 50 #F8FAFC
        private readonly Color _subTextColor = Color.FromArgb(148, 163, 184);// Slate 400 #94A3B8

        // Controls
        private Panel _pnlHeader = null!;
        private Label _lblTitle = null!;
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

            // 1. HEADER TOOLBAR PANEL (DOCK TOP)
            _pnlHeader = new Panel
            {
                Dock = DockStyle.Top,
                Height = 38,
                BackColor = _cardColor,
                Padding = new Padding(8, 4, 8, 4)
            };

            _btnRefresh = new Button
            {
                Dock = DockStyle.Right,
                Width = 80,
                Text = "Refresh",
                Font = new Font("Segoe UI", 8.5f, FontStyle.Bold),
                ForeColor = Color.White,
                BackColor = Color.FromArgb(59, 130, 246),
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            _btnRefresh.FlatAppearance.BorderSize = 0;
            _btnRefresh.Click += (s, e) => RefreshData();

            _lblTitle = new Label
            {
                Dock = DockStyle.Left,
                AutoSize = true,
                Text = "NAVIGATION",
                Font = new Font("Segoe UI", 9.5f, FontStyle.Bold),
                ForeColor = Color.FromArgb(56, 189, 248),
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(0, 4, 0, 0)
            };

            _pnlHeader.Controls.Add(_btnRefresh);
            _pnlHeader.Controls.Add(_lblTitle);

            // 2. STATUS LABEL (FOOTER - DOCK BOTTOM)
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

            // 3. SPLIT CONTAINER (VÙNG 1 & VÙNG 2 - DOCK FILL)
            _splitContainer = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Horizontal,
                SplitterWidth = 8,
                Panel1MinSize = 80,
                Panel2MinSize = 80,
                BackColor = Color.FromArgb(51, 65, 85) // Slate 700 visible draggable splitter bar
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
                Text = "OPEN FILES",
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
                ItemHeight = 36
            };
            _lstWorkbooks.DrawItem += DrawWorkbookItem;
            _lstWorkbooks.SelectedIndexChanged += LstWorkbooks_SelectedIndexChanged;

            // CORRECT DOCK ORDER IN PANEL 1: TOP HEADER FIRST, FILL LISTBOX SECOND
            _splitContainer.Panel1.Controls.Add(_pnlVung1Header);
            _splitContainer.Panel1.Controls.Add(_lstWorkbooks);
            _splitContainer.Panel1.BackColor = _bgColor;

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
                Text = "SHEETS",
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
                ItemHeight = 36
            };
            _lstSheets.DrawItem += DrawSheetItem;
            _lstSheets.SelectedIndexChanged += LstSheets_SelectedIndexChanged;

            // CORRECT DOCK ORDER IN PANEL 2: TOP HEADER FIRST, FILL LISTBOX SECOND
            _splitContainer.Panel2.Controls.Add(_pnlVung2Header);
            _splitContainer.Panel2.Controls.Add(_lstSheets);
            _splitContainer.Panel2.BackColor = _bgColor;

            // ADD CONTROLS IN PROPER WINFORMS DOCK ORDER: TOP -> BOTTOM -> FILL
            this.Controls.Add(_pnlHeader);
            this.Controls.Add(_lblStatus);
            this.Controls.Add(_splitContainer);

            this.ResumeLayout(false);
        }

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            SetEqualSplitterDistance();
            RefreshData();
        }

        private void SetEqualSplitterDistance()
        {
            try
            {
                if (_splitContainer != null && this.Height > 160)
                {
                    int availableHeight = this.Height - (_pnlHeader?.Height ?? 38) - (_lblStatus?.Height ?? 24);
                    int half = availableHeight / 2;
                    if (half >= _splitContainer.Panel1MinSize && half <= availableHeight - _splitContainer.Panel2MinSize)
                    {
                        _splitContainer.SplitterDistance = half;
                    }
                }
            }
            catch { }
        }

        public void SelectWorkbookByName(string wbName)
        {
            if (string.IsNullOrEmpty(wbName)) return;

            var match = _allWorkbooks.FirstOrDefault(w => w.Name.Equals(wbName, StringComparison.OrdinalIgnoreCase));
            if (match != null)
            {
                _isUpdatingUi = true;
                SelectWorkbookItemInList(match);
                _isUpdatingUi = false;
            }
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

                _lblVung1Title.Text = $"OPEN FILES ({_allWorkbooks.Count})";
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
                    _lblVung2Title.Text = "SHEETS";
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

                _lblVung2Title.Text = $"SHEETS ({_allSheets.Count})";
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
                        try
                        {
                            BringWindowToTop(hwnd);
                            SetForegroundWindow(hwnd);
                        }
                        catch { }
                    }
                }

                // Synchronize selection across all TaskPanes so File Tháng 06's TaskPane highlights File Tháng 06!
                AddIn.SyncSelectedWorkbookInAllTaskPanes(item.Name);

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
                            try
                            {
                                BringWindowToTop(hwnd);
                                SetForegroundWindow(hwnd);
                            }
                            catch { }
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
                            try
                            {
                                BringWindowToTop(hwnd);
                                SetForegroundWindow(hwnd);
                            }
                            catch { }
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
            TextRenderer.DrawText(e.Graphics, "📄", new Font("Segoe UI", 10.5f), new Point(e.Bounds.X + 6, e.Bounds.Y + 7), _textColor);

            // Draw File Name across full width with EndEllipsis
            using (var fontName = new Font("Segoe UI", 9.5f, FontStyle.Bold))
            {
                Rectangle nameRect = new Rectangle(e.Bounds.X + 30, e.Bounds.Y + 4, e.Bounds.Width - 36, e.Bounds.Height - 8);
                TextRenderer.DrawText(e.Graphics, item.Name, fontName, nameRect, _textColor, TextFormatFlags.VerticalCenter | TextFormatFlags.Left | TextFormatFlags.EndEllipsis);
            }
        }

        private void DrawSheetItem(object? sender, DrawItemEventArgs e)
        {
            if (e.Index < 0 || e.Index >= _lstSheets.Items.Count) return;
            var item = (SheetItem)_lstSheets.Items[e.Index];

            bool isSelected = (e.Index == _lstSheets.SelectedIndex);
            Color tabColor = ColorTranslator.FromHtml(item.TabColorHex);

            Color itemBg;
            Color textCol;

            if (isSelected)
            {
                itemBg = Color.FromArgb(79, 70, 229); // Indigo 600 if selected
                textCol = Color.White;
            }
            else if (item.HasCustomTabColor)
            {
                itemBg = tabColor;
                // High contrast text color calculation
                double brightness = (tabColor.R * 0.299 + tabColor.G * 0.587 + tabColor.B * 0.114);
                textCol = brightness > 160 ? Color.FromArgb(15, 23, 42) : Color.White;
            }
            else
            {
                itemBg = _cardColor;
                textCol = _textColor;
            }

            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;

            // Draw Card Background
            using (var brush = new SolidBrush(itemBg))
            {
                e.Graphics.FillRectangle(brush, new Rectangle(e.Bounds.X + 2, e.Bounds.Y + 2, e.Bounds.Width - 4, e.Bounds.Height - 4));
            }

            // Draw subtle vertical tab color strip only for non-custom cards
            if (!item.HasCustomTabColor && !isSelected)
            {
                using (var tabBrush = new SolidBrush(tabColor))
                {
                    e.Graphics.FillRectangle(tabBrush, new Rectangle(e.Bounds.X + 6, e.Bounds.Y + 6, 5, e.Bounds.Height - 12));
                }
            }

            // Draw Sheet Name across full width
            int leftOffset = (item.HasCustomTabColor || isSelected) ? 10 : 18;
            int rightPadding = item.IsVisible ? 10 : 56;
            using (var fontName = new Font("Segoe UI", 9.5f, FontStyle.Bold))
            {
                Rectangle nameRect = new Rectangle(e.Bounds.X + leftOffset, e.Bounds.Y + 4, e.Bounds.Width - leftOffset - rightPadding, e.Bounds.Height - 8);
                TextRenderer.DrawText(e.Graphics, item.Name, fontName, nameRect, textCol, TextFormatFlags.VerticalCenter | TextFormatFlags.Left | TextFormatFlags.EndEllipsis);
            }

            // Draw Hidden Badge if hidden
            if (!item.IsVisible)
            {
                using (var fontBadge = new Font("Segoe UI", 7.5f, FontStyle.Bold))
                using (var badgeBg = new SolidBrush(Color.FromArgb(239, 68, 68))) // Red 500
                {
                    Rectangle badgeRect = new Rectangle(e.Bounds.Right - 52, e.Bounds.Y + 8, 46, 18);
                    e.Graphics.FillRectangle(badgeBg, badgeRect);
                    TextRenderer.DrawText(e.Graphics, "Hidden", fontBadge, new Point(badgeRect.X + 5, badgeRect.Y + 2), Color.White);
                }
            }
        }
    }
}
