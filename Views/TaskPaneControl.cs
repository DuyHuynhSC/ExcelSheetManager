using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
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

        // Theme Flag & Colors
        private bool _isDarkTheme = true;
        private bool _showHiddenSheetsInList = true; // Toggle for Vung 2 filter list

        private Color _bgColor;
        private Color _cardColor;
        private Color _selectColor;
        private Color _textColor;
        private Color _subTextColor;
        private Color _splitterColor;

        // Layout Containers
        private TableLayoutPanel _mainTable = null!;
        private TableLayoutPanel _tblVung1 = null!;
        private TableLayoutPanel _tblVung2 = null!;

        // Controls
        private SplitContainer _splitContainer = null!;

        private Panel _pnlVung1Header = null!;
        private Panel _pnlVung1Top = null!;
        private Label _lblVung1Title = null!;
        private Button _btnRefresh = null!;
        private Button _btnTheme = null!;
        private Button _btnAi = null!;
        private TextBox _txtFilterFile = null!;
        private ListBox _lstWorkbooks = null!;

        private Panel _pnlVung2Header = null!;
        private Panel _pnlVung2Top = null!;
        private Label _lblVung2Title = null!;
        private Button _btnToc = null!;
        private Button _btnToggleHidden = null!;
        private Button _btnCopySheetName = null!;
        private TextBox _txtFilterSheet = null!;
        private ListBox _lstSheets = null!;

        // Context Menu
        private ContextMenuStrip _cmsSheetMenu = null!;
        private ToolStripMenuItem _miProtectSheet = null!;
        private ToolStripMenuItem _miCopyName = null!;
        private ToolStripMenuItem _miToggleHideExcel = null!;
        private ToolStripMenuItem _miChangeTabColor = null!;
        private ToolStripMenuItem _miClearTabColor = null!;
        private ToolStripMenuItem _miSortMenu = null!;
        private ToolStripMenuItem _miSortAZ = null!;
        private ToolStripMenuItem _miSortZA = null!;
        private ToolStripMenuItem _miSortColor = null!;
        private ToolStripMenuItem _miExportMenu = null!;
        private ToolStripMenuItem _miExportPdf = null!;
        private ToolStripMenuItem _miExportExcel = null!;

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
            this.Font = new Font("Segoe UI", 9.5f, FontStyle.Regular);

            _isDarkTheme = SettingsHelper.GetIsDarkTheme(defaultValue: true);

            InitializeComponent();
            ApplyTheme();
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();

            // 1. ROOT TABLE LAYOUT PANEL (2 ROWS: CONTENT 100%, FOOTER 24px)
            // Title moved to Excel's native TaskPane Header to maximize vertical space for Vung 1 and Vung 2!
            _mainTable = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 2,
                Margin = new Padding(0),
                Padding = new Padding(0)
            };
            _mainTable.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
            _mainTable.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));  // Row 0: SplitContainer
            _mainTable.RowStyles.Add(new RowStyle(SizeType.Absolute, 24f)); // Row 1: Status Footer

            // 2. STATUS LABEL (ROW 1 - FOOTER)
            _lblStatus = new Label
            {
                Dock = DockStyle.Fill,
                Margin = new Padding(0),
                Font = new Font("Segoe UI", 8.5f),
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(8, 0, 0, 0),
                Text = "Ready"
            };
            _mainTable.Controls.Add(_lblStatus, 0, 1);

            // 3. SPLIT CONTAINER (ROW 0 - CONTENT)
            _splitContainer = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Horizontal,
                SplitterWidth = 8,
                Panel1MinSize = 80,
                Panel2MinSize = 80,
                Margin = new Padding(0)
            };

            // --- VÙNG 1: TABLE LAYOUT (ROW 0: HEADER 55px, ROW 1: LISTBOX 100%) ---
            _tblVung1 = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 2,
                Margin = new Padding(0),
                Padding = new Padding(0)
            };
            _tblVung1.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
            _tblVung1.RowStyles.Add(new RowStyle(SizeType.Absolute, 55f));
            _tblVung1.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));

            _pnlVung1Header = new Panel
            {
                Dock = DockStyle.Fill,
                Margin = new Padding(0),
                Padding = new Padding(6, 4, 6, 4)
            };

            _pnlVung1Top = new Panel
            {
                Dock = DockStyle.Top,
                Height = 24,
                Margin = new Padding(0)
            };

            _lblVung1Title = new Label
            {
                Text = "OPEN FILES",
                Font = new Font("Segoe UI", 9.5f, FontStyle.Bold),
                ForeColor = Color.FromArgb(56, 189, 248),
                Dock = DockStyle.Left,
                AutoSize = true,
                TextAlign = ContentAlignment.MiddleLeft
            };

            // AI Assistant Button (Far Right of Vung 1 - Rose Pink #EC4899)
            _btnAi = new Button
            {
                Dock = DockStyle.Right,
                Width = 36,
                Text = "AI",
                Font = new Font("Segoe UI", 8f, FontStyle.Bold),
                ForeColor = Color.White,
                BackColor = Color.FromArgb(236, 72, 153),
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            _btnAi.FlatAppearance.BorderSize = 0;
            _btnAi.Click += (s, e) => AiAssistantForm.ShowForm();

            // 4px Gap
            Panel spacerV1_1 = new Panel { Dock = DockStyle.Right, Width = 4, BackColor = Color.Transparent };

            // Theme Toggle Button (Middle Right of Vung 1)
            _btnTheme = new Button
            {
                Dock = DockStyle.Right,
                Width = 52,
                Text = _isDarkTheme ? "Light" : "Dark",
                Font = new Font("Segoe UI", 8f, FontStyle.Bold),
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            _btnTheme.FlatAppearance.BorderSize = 0;
            _btnTheme.Click += (s, e) => ToggleTheme();

            // 4px Gap
            Panel spacerV1_2 = new Panel { Dock = DockStyle.Right, Width = 4, BackColor = Color.Transparent };

            // Refresh Button (Left of Theme in Vung 1 - Blue #3B82F6)
            _btnRefresh = new Button
            {
                Dock = DockStyle.Right,
                Width = 58,
                Text = "Refresh",
                Font = new Font("Segoe UI", 8f, FontStyle.Bold),
                ForeColor = Color.White,
                BackColor = Color.FromArgb(59, 130, 246),
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            _btnRefresh.FlatAppearance.BorderSize = 0;
            _btnRefresh.Click += (s, e) => RefreshData();

            _pnlVung1Top.Controls.Add(_lblVung1Title);
            _pnlVung1Top.Controls.Add(_btnAi);         // Far Right (AI)
            _pnlVung1Top.Controls.Add(spacerV1_1);      // Gap
            _pnlVung1Top.Controls.Add(_btnTheme);       // Middle (Theme)
            _pnlVung1Top.Controls.Add(spacerV1_2);      // Gap
            _pnlVung1Top.Controls.Add(_btnRefresh);     // Left (Refresh)

            _txtFilterFile = new TextBox
            {
                Dock = DockStyle.Bottom,
                BorderStyle = BorderStyle.FixedSingle,
                Font = new Font("Segoe UI", 9f),
                Text = ""
            };
            _txtFilterFile.TextChanged += (s, e) => FilterWorkbooksList();

            _pnlVung1Header.Controls.Add(_pnlVung1Top);
            _pnlVung1Header.Controls.Add(_txtFilterFile);

            _lstWorkbooks = new ListBox
            {
                Dock = DockStyle.Fill,
                Margin = new Padding(0),
                BorderStyle = BorderStyle.None,
                DrawMode = DrawMode.OwnerDrawFixed,
                ItemHeight = 36
            };
            _lstWorkbooks.DrawItem += DrawWorkbookItem;
            _lstWorkbooks.SelectedIndexChanged += LstWorkbooks_SelectedIndexChanged;

            _tblVung1.Controls.Add(_pnlVung1Header, 0, 0);
            _tblVung1.Controls.Add(_lstWorkbooks, 0, 1);
            _splitContainer.Panel1.Controls.Add(_tblVung1);

            // --- VÙNG 2: TABLE LAYOUT (ROW 0: HEADER 55px, ROW 1: LISTBOX 100%) ---
            _tblVung2 = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 2,
                Margin = new Padding(0),
                Padding = new Padding(0)
            };
            _tblVung2.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
            _tblVung2.RowStyles.Add(new RowStyle(SizeType.Absolute, 55f));
            _tblVung2.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));

            _pnlVung2Header = new Panel
            {
                Dock = DockStyle.Fill,
                Margin = new Padding(0),
                Padding = new Padding(6, 4, 6, 4)
            };

            _pnlVung2Top = new Panel
            {
                Dock = DockStyle.Top,
                Height = 24,
                Margin = new Padding(0)
            };

            _lblVung2Title = new Label
            {
                Text = "SHEETS",
                Font = new Font("Segoe UI", 9.5f, FontStyle.Bold),
                ForeColor = Color.FromArgb(129, 140, 248),
                Dock = DockStyle.Left,
                AutoSize = true,
                TextAlign = ContentAlignment.MiddleLeft
            };

            // Copy Sheet Name Button (Far Right - Emerald Green #16A34A)
            _btnCopySheetName = new Button
            {
                Dock = DockStyle.Right,
                Width = 48,
                Text = "Copy",
                Font = new Font("Segoe UI", 8f, FontStyle.Bold),
                ForeColor = Color.White,
                BackColor = Color.FromArgb(16, 185, 129), // Emerald Green
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            _btnCopySheetName.FlatAppearance.BorderSize = 0;
            _btnCopySheetName.Click += (s, e) => CopySelectedSheetName();

            // 4px Physical Gap 1
            Panel spacerV2_1 = new Panel { Dock = DockStyle.Right, Width = 4, BackColor = Color.Transparent };

            // Hide/Show Filter Button (Toggles List View Filtering of Hidden Sheets - Sky Blue #0284C7)
            _btnToggleHidden = new Button
            {
                Dock = DockStyle.Right,
                Width = 78,
                Text = "Hide/Show",
                Font = new Font("Segoe UI", 8f, FontStyle.Bold),
                ForeColor = Color.White,
                BackColor = Color.FromArgb(2, 132, 199), // Sky Blue
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            _btnToggleHidden.FlatAppearance.BorderSize = 0;
            _btnToggleHidden.Click += (s, e) => ToggleHiddenSheetsListFilter();

            // 4px Physical Gap 2
            Panel spacerV2_2 = new Panel { Dock = DockStyle.Right, Width = 4, BackColor = Color.Transparent };

            // Table of Contents Generator Button (Feature 2 - Purple #8B5CF6)
            _btnToc = new Button
            {
                Dock = DockStyle.Right,
                Width = 45,
                Text = "TOC",
                Font = new Font("Segoe UI", 8f, FontStyle.Bold),
                ForeColor = Color.White,
                BackColor = Color.FromArgb(139, 92, 246), // Purple
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            _btnToc.FlatAppearance.BorderSize = 0;
            _btnToc.Click += (s, e) => GenerateTableOfContents();

            _pnlVung2Top.Controls.Add(_lblVung2Title);
            _pnlVung2Top.Controls.Add(_btnCopySheetName);  // Far Right (Green)
            _pnlVung2Top.Controls.Add(spacerV2_1);          // Gap
            _pnlVung2Top.Controls.Add(_btnToggleHidden);   // Middle (Blue)
            _pnlVung2Top.Controls.Add(spacerV2_2);          // Gap
            _pnlVung2Top.Controls.Add(_btnToc);            // Left of Hide/Show (Purple)

            _txtFilterSheet = new TextBox
            {
                Dock = DockStyle.Bottom,
                BorderStyle = BorderStyle.FixedSingle,
                Font = new Font("Segoe UI", 9f),
                Text = ""
            };
            _txtFilterSheet.TextChanged += (s, e) => FilterSheetsList();

            _pnlVung2Header.Controls.Add(_pnlVung2Top);
            _pnlVung2Header.Controls.Add(_txtFilterSheet);

            _lstSheets = new ListBox
            {
                Dock = DockStyle.Fill,
                Margin = new Padding(0),
                BorderStyle = BorderStyle.None,
                DrawMode = DrawMode.OwnerDrawFixed,
                ItemHeight = 36
            };
            _lstSheets.DrawItem += DrawSheetItem;
            _lstSheets.SelectedIndexChanged += LstSheets_SelectedIndexChanged;
            _lstSheets.MouseDown += LstSheets_MouseDown;

            // Context Menu Initialization
            _cmsSheetMenu = new ContextMenuStrip();
            
            _miProtectSheet = new ToolStripMenuItem("🔒 Protect Sheet...");
            _miProtectSheet.Click += (s, e) => ToggleProtectSelectedSheet();

            _miCopyName = new ToolStripMenuItem("📋 Copy Sheet Name");
            _miCopyName.Click += (s, e) => CopySelectedSheetName();

            _miToggleHideExcel = new ToolStripMenuItem("👁️ Hide / Unhide Sheet in Excel");
            _miToggleHideExcel.Click += (s, e) => ToggleSheetVisibilityInExcel();

            _miChangeTabColor = new ToolStripMenuItem("🎨 Change Tab Color...");
            _miChangeTabColor.Click += (s, e) => ChangeSheetTabColor();

            _miClearTabColor = new ToolStripMenuItem("🧹 Reset Tab Color");
            _miClearTabColor.Click += (s, e) => ResetSheetTabColor();

            // Submenu: Sort Sheets
            _miSortMenu = new ToolStripMenuItem("🔤 Sort Sheets in Excel");
            _miSortAZ = new ToolStripMenuItem("Sort A to Z", null, (s, e) => SortSheetsInExcel("AZ"));
            _miSortZA = new ToolStripMenuItem("Sort Z to A", null, (s, e) => SortSheetsInExcel("ZA"));
            _miSortColor = new ToolStripMenuItem("Sort by Tab Color", null, (s, e) => SortSheetsInExcel("COLOR"));
            _miSortMenu.DropDownItems.Add(_miSortAZ);
            _miSortMenu.DropDownItems.Add(_miSortZA);
            _miSortMenu.DropDownItems.Add(_miSortColor);

            // Submenu: Export Sheet
            _miExportMenu = new ToolStripMenuItem("📤 Export Sheet");
            _miExportPdf = new ToolStripMenuItem("📄 Export to PDF...", null, (s, e) => ExportSheetToPdf());
            _miExportExcel = new ToolStripMenuItem("📁 Export to New Excel File...", null, (s, e) => ExportSheetToNewWorkbook());
            _miExportMenu.DropDownItems.Add(_miExportPdf);
            _miExportMenu.DropDownItems.Add(_miExportExcel);

            _cmsSheetMenu.Items.Add(_miProtectSheet);
            _cmsSheetMenu.Items.Add(_miCopyName);
            _cmsSheetMenu.Items.Add(_miToggleHideExcel);
            _cmsSheetMenu.Items.Add(new ToolStripSeparator());
            _cmsSheetMenu.Items.Add(_miChangeTabColor);
            _cmsSheetMenu.Items.Add(_miClearTabColor);
            _cmsSheetMenu.Items.Add(new ToolStripSeparator());
            _cmsSheetMenu.Items.Add(_miSortMenu);
            _cmsSheetMenu.Items.Add(_miExportMenu);

            _tblVung2.Controls.Add(_pnlVung2Header, 0, 0);
            _tblVung2.Controls.Add(_lstSheets, 0, 1);
            _splitContainer.Panel2.Controls.Add(_tblVung2);

            _mainTable.Controls.Add(_splitContainer, 0, 0);

            // ADD ROOT TABLE TO CONTROL
            this.Controls.Add(_mainTable);

            this.ResumeLayout(false);
        }

        private void LstSheets_MouseDown(object? sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                int index = _lstSheets.IndexFromPoint(e.Location);
                if (index >= 0 && index < _lstSheets.Items.Count)
                {
                    _lstSheets.SelectedIndex = index;
                    if (_lstSheets.Items[index] is SheetItem item)
                    {
                        _miProtectSheet.Text = item.IsProtected ? "🔓 Unprotect Sheet..." : "🔒 Protect Sheet...";
                        _miToggleHideExcel.Text = item.IsVisible ? "👁️ Hide Sheet in Excel" : "👁️ Unhide Sheet in Excel";
                        _cmsSheetMenu.Show(_lstSheets, e.Location);
                    }
                }
            }
        }

        private void ToggleTheme()
        {
            _isDarkTheme = !_isDarkTheme;
            SettingsHelper.SetIsDarkTheme(_isDarkTheme);
            ApplyTheme();
        }

        private void ApplyTheme()
        {
            _bgColor = _isDarkTheme ? Color.FromArgb(15, 23, 42) : Color.FromArgb(241, 245, 249);
            _cardColor = _isDarkTheme ? Color.FromArgb(30, 41, 59) : Color.FromArgb(255, 255, 255);
            _selectColor = Color.FromArgb(2, 132, 199);
            _textColor = _isDarkTheme ? Color.FromArgb(248, 250, 252) : Color.FromArgb(15, 23, 42);
            _subTextColor = _isDarkTheme ? Color.FromArgb(148, 163, 184) : Color.FromArgb(100, 116, 139);
            _splitterColor = _isDarkTheme ? Color.FromArgb(51, 65, 85) : Color.FromArgb(203, 213, 225);

            this.BackColor = _bgColor;
            this.ForeColor = _textColor;

            _mainTable.BackColor = _bgColor;

            _btnTheme.Text = _isDarkTheme ? "Light" : "Dark";
            _btnTheme.BackColor = _isDarkTheme ? Color.FromArgb(51, 65, 85) : Color.FromArgb(30, 41, 59);
            _btnTheme.ForeColor = Color.White;

            _btnRefresh.BackColor = Color.FromArgb(59, 130, 246); // Sky Blue
            _btnRefresh.ForeColor = Color.White;

            _btnAi.BackColor = Color.FromArgb(236, 72, 153); // Rose Pink
            _btnAi.ForeColor = Color.White;

            _btnToggleHidden.BackColor = Color.FromArgb(2, 132, 199);  // Sky Blue
            _btnToggleHidden.ForeColor = Color.White;

            _btnCopySheetName.BackColor = Color.FromArgb(16, 185, 129); // Emerald Green
            _btnCopySheetName.ForeColor = Color.White;

            _btnToc.BackColor = Color.FromArgb(139, 92, 246); // Purple
            _btnToc.ForeColor = Color.White;

            _lblStatus.BackColor = _cardColor;
            _lblStatus.ForeColor = _subTextColor;

            _splitContainer.BackColor = _splitterColor;

            _tblVung1.BackColor = _bgColor;
            _pnlVung1Header.BackColor = _bgColor;
            _txtFilterFile.BackColor = _cardColor;
            _txtFilterFile.ForeColor = _textColor;
            _lstWorkbooks.BackColor = _bgColor;
            _lstWorkbooks.ForeColor = _textColor;

            _tblVung2.BackColor = _bgColor;
            _pnlVung2Header.BackColor = _bgColor;
            _txtFilterSheet.BackColor = _cardColor;
            _txtFilterSheet.ForeColor = _textColor;
            _lstSheets.BackColor = _bgColor;
            _lstSheets.ForeColor = _textColor;

            _lstWorkbooks.Invalidate();
            _lstSheets.Invalidate();
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
                if (_splitContainer != null && _splitContainer.Height > 160)
                {
                    int half = _splitContainer.Height / 2;
                    if (half >= _splitContainer.Panel1MinSize && half <= _splitContainer.Height - _splitContainer.Panel2MinSize)
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
                            IsProtected = ws.ProtectContents,
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
                            IsProtected = false,
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

                var matches = _allSheets.AsEnumerable();

                // User Directive: Toggle filtering of hidden sheets in Vung 2 list without changing Excel worksheet state!
                if (!_showHiddenSheetsInList)
                {
                    matches = matches.Where(s => s.IsVisible);
                }

                string filter = _txtFilterSheet.Text.Trim();
                if (!string.IsNullOrEmpty(filter))
                {
                    matches = matches.Where(s => s.Name.IndexOf(filter, StringComparison.OrdinalIgnoreCase) >= 0);
                }

                foreach (var item in matches)
                {
                    _lstSheets.Items.Add(item);
                }

                _lblVung2Title.Text = $"SHEETS ({_lstSheets.Items.Count})";
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

        // --- BUTTON & CONTEXT ACTION HANDLERS ---
        private void ToggleHiddenSheetsListFilter()
        {
            _showHiddenSheetsInList = !_showHiddenSheetsInList;
            FilterSheetsList();
            _lblStatus.Text = _showHiddenSheetsInList ? "Showing all sheets (including hidden)" : "Filtering out hidden sheets from list";
        }

        private void ToggleSheetVisibilityInExcel()
        {
            try
            {
                if (_selectedWorkbook?.WorkbookRef is not Excel.Workbook wb) return;

                if (_lstSheets.SelectedIndex >= 0 && _lstSheets.SelectedIndex < _lstSheets.Items.Count)
                {
                    if (_lstSheets.Items[_lstSheets.SelectedIndex] is SheetItem targetSheet && targetSheet.SheetRef is Excel.Worksheet ws)
                    {
                        if (ws.Visible == Excel.XlSheetVisibility.xlSheetVisible)
                        {
                            int visibleCount = 0;
                            foreach (Excel.Worksheet sheet in wb.Worksheets)
                            {
                                if (sheet.Visible == Excel.XlSheetVisibility.xlSheetVisible) visibleCount++;
                            }

                            if (visibleCount <= 1)
                            {
                                _lblStatus.Text = "Cannot hide the only visible sheet in workbook.";
                                return;
                            }

                            ws.Visible = Excel.XlSheetVisibility.xlSheetHidden;
                            _lblStatus.Text = $"Hid sheet in Excel: {ws.Name}";
                        }
                        else
                        {
                            ws.Visible = Excel.XlSheetVisibility.xlSheetVisible;
                            _lblStatus.Text = $"Unhid sheet in Excel: {ws.Name}";
                        }

                        RefreshData();
                    }
                }
            }
            catch (Exception ex)
            {
                _lblStatus.Text = $"Could not toggle sheet visibility: {ex.Message}";
            }
        }

        private void CopySelectedSheetName()
        {
            try
            {
                string? nameToCopy = null;

                if (_lstSheets.SelectedIndex >= 0 && _lstSheets.SelectedIndex < _lstSheets.Items.Count)
                {
                    if (_lstSheets.Items[_lstSheets.SelectedIndex] is SheetItem item)
                    {
                        nameToCopy = item.Name;
                    }
                }

                if (string.IsNullOrEmpty(nameToCopy) && _selectedWorkbook?.WorkbookRef is Excel.Workbook refWb)
                {
                    object activeSheet = refWb.ActiveSheet;
                    if (activeSheet is Excel.Worksheet ws) nameToCopy = ws.Name;
                    else if (activeSheet is Excel.Chart chart) nameToCopy = chart.Name;
                }

                if (!string.IsNullOrEmpty(nameToCopy))
                {
                    Clipboard.SetText(nameToCopy);
                    _lblStatus.Text = $"Copied sheet name: \"{nameToCopy}\" to clipboard";
                }
                else
                {
                    _lblStatus.Text = "No sheet selected to copy.";
                }
            }
            catch (Exception ex)
            {
                _lblStatus.Text = $"Could not copy sheet name: {ex.Message}";
            }
        }

        private void ToggleProtectSelectedSheet()
        {
            try
            {
                if (_lstSheets.SelectedIndex >= 0 && _lstSheets.SelectedIndex < _lstSheets.Items.Count)
                {
                    if (_lstSheets.Items[_lstSheets.SelectedIndex] is SheetItem targetSheet && targetSheet.SheetRef is Excel.Worksheet ws)
                    {
                        var excelApp = (Excel.Application)ExcelDnaUtil.Application;

                        if (ws.ProtectContents)
                        {
                            try
                            {
                                ws.Unprotect();
                                _lblStatus.Text = $"Unprotected sheet: {ws.Name}";
                            }
                            catch
                            {
                                if (excelApp != null)
                                {
                                    object input = excelApp.InputBox($"Enter password to unprotect sheet '{ws.Name}':", "Unprotect Sheet", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 2);
                                    if (input is string pwd && !string.IsNullOrEmpty(pwd))
                                    {
                                        ws.Unprotect(pwd);
                                        _lblStatus.Text = $"Unprotected sheet: {ws.Name}";
                                    }
                                }
                            }
                        }
                        else
                        {
                            if (excelApp != null)
                            {
                                object input = excelApp.InputBox($"Enter password to protect sheet '{ws.Name}' (or click OK with blank):", "Protect Sheet", "", Type.Missing, Type.Missing, Type.Missing, Type.Missing, 2);
                                if (input is bool b && !b) return; // User clicked Cancel

                                string pwd = input as string ?? string.Empty;
                                if (string.IsNullOrEmpty(pwd))
                                {
                                    ws.Protect();
                                }
                                else
                                {
                                    ws.Protect(pwd);
                                }
                                _lblStatus.Text = $"Protected sheet: {ws.Name}";
                            }
                            else
                            {
                                ws.Protect();
                                _lblStatus.Text = $"Protected sheet: {ws.Name}";
                            }
                        }

                        RefreshData();
                    }
                }
            }
            catch (Exception ex)
            {
                _lblStatus.Text = $"Protection error: {ex.Message}";
            }
        }

        // --- FEATURE 2: TABLE OF CONTENTS GENERATOR ---
        private void GenerateTableOfContents()
        {
            try
            {
                if (_selectedWorkbook?.WorkbookRef is not Excel.Workbook wb)
                {
                    _lblStatus.Text = "No active workbook to create Table of Contents.";
                    return;
                }

                const string tocSheetName = "MUC_LUC";
                Excel.Worksheet? tocWs = null;

                foreach (Excel.Worksheet sheet in wb.Worksheets)
                {
                    if (sheet.Name.Equals(tocSheetName, StringComparison.OrdinalIgnoreCase))
                    {
                        tocWs = sheet;
                        break;
                    }
                }

                if (tocWs == null)
                {
                    Excel.Worksheet firstWs = (Excel.Worksheet)wb.Worksheets[1];
                    tocWs = (Excel.Worksheet)wb.Worksheets.Add(firstWs, Type.Missing, Type.Missing, Type.Missing);
                    tocWs.Name = tocSheetName;
                }

                tocWs.Activate();
                tocWs.Cells.Clear();

                Excel.Range titleRange = tocWs.get_Range("A1", "E1");
                titleRange.Merge();
                titleRange.Value2 = "📑 BẢNG MỤC LỤC SỔ TÍNH (TABLE OF CONTENTS)";
                titleRange.Font.Bold = true;
                titleRange.Font.Size = 14;
                titleRange.Font.Name = "Segoe UI";
                titleRange.Interior.Color = ColorTranslator.ToOle(Color.FromArgb(30, 41, 59));
                titleRange.Font.Color = ColorTranslator.ToOle(Color.FromArgb(248, 250, 252));
                titleRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                titleRange.RowHeight = 32;

                string[] headers = { "STT", "TÊN SHEET (CLICK NHẢY TRANG)", "MÀU TAB", "TRẠNG THÁI", "BẢO VỆ" };
                for (int c = 0; c < headers.Length; c++)
                {
                    Excel.Range cell = (Excel.Range)tocWs.Cells[3, c + 1];
                    cell.Value2 = headers[c];
                    cell.Font.Bold = true;
                    cell.Font.Size = 10;
                    cell.Font.Name = "Segoe UI";
                    cell.Interior.Color = ColorTranslator.ToOle(Color.FromArgb(51, 65, 85));
                    cell.Font.Color = ColorTranslator.ToOle(Color.White);
                    cell.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                }
                ((Excel.Range)tocWs.get_Range("A3", "E3")).RowHeight = 24;

                int rowIndex = 4;
                int stt = 1;
                foreach (Excel.Worksheet sheet in wb.Worksheets)
                {
                    if (sheet.Name.Equals(tocSheetName, StringComparison.OrdinalIgnoreCase)) continue;

                    Excel.Range cellStt = (Excel.Range)tocWs.Cells[rowIndex, 1];
                    cellStt.Value2 = stt++;
                    cellStt.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                    Excel.Range cellName = (Excel.Range)tocWs.Cells[rowIndex, 2];
                    tocWs.Hyperlinks.Add(cellName, "", $"'{sheet.Name}'!A1", Type.Missing, sheet.Name);
                    cellName.Font.Bold = true;
                    cellName.Font.Size = 10;

                    Excel.Range cellColor = (Excel.Range)tocWs.Cells[rowIndex, 3];
                    var (colorHex, hasCustom) = ColorHelper.GetSheetTabColorHex(sheet);
                    cellColor.Value2 = colorHex;
                    cellColor.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    if (hasCustom)
                    {
                        try
                        {
                            Color c = ColorTranslator.FromHtml(colorHex);
                            cellColor.Interior.Color = ColorTranslator.ToOle(c);
                        }
                        catch { }
                    }

                    Excel.Range cellVis = (Excel.Range)tocWs.Cells[rowIndex, 4];
                    cellVis.Value2 = sheet.Visible == Excel.XlSheetVisibility.xlSheetVisible ? "Hiện" : "Ẩn";
                    cellVis.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                    Excel.Range cellProt = (Excel.Range)tocWs.Cells[rowIndex, 5];
                    cellProt.Value2 = sheet.ProtectContents ? "🔒 Khóa" : "Bình thường";
                    cellProt.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                    ((Excel.Range)tocWs.get_Range($"A{rowIndex}", $"E{rowIndex}")).RowHeight = 20;
                    rowIndex++;
                }

                Excel.Range fullTable = tocWs.get_Range("A3", $"E{rowIndex - 1}");
                fullTable.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                fullTable.Columns.AutoFit();

                RefreshData();
                _lblStatus.Text = $"Created Table of Contents 'MUC_LUC' with {stt - 1} sheets";
            }
            catch (Exception ex)
            {
                _lblStatus.Text = $"Could not create TOC: {ex.Message}";
            }
        }

        // --- FEATURE 3: SHEET SORTING ---
        private void SortSheetsInExcel(string sortMode)
        {
            try
            {
                if (_selectedWorkbook?.WorkbookRef is not Excel.Workbook wb) return;

                List<Excel.Worksheet> sheetsList = new();
                foreach (object s in wb.Worksheets)
                {
                    if (s is Excel.Worksheet ws) sheetsList.Add(ws);
                }

                if (sheetsList.Count <= 1) return;

                if (sortMode == "AZ")
                {
                    sheetsList = sheetsList.OrderBy(s => s.Name, StringComparer.OrdinalIgnoreCase).ToList();
                }
                else if (sortMode == "ZA")
                {
                    sheetsList = sheetsList.OrderByDescending(s => s.Name, StringComparer.OrdinalIgnoreCase).ToList();
                }
                else if (sortMode == "COLOR")
                {
                    sheetsList = sheetsList.OrderBy(s => ColorHelper.GetSheetTabColorHex(s).Hex).ToList();
                }

                for (int i = 0; i < sheetsList.Count; i++)
                {
                    if (i == 0)
                    {
                        sheetsList[i].Move(wb.Worksheets[1], Type.Missing);
                    }
                    else
                    {
                        sheetsList[i].Move(Type.Missing, sheetsList[i - 1]);
                    }
                }

                RefreshData();
                _lblStatus.Text = $"Sorted sheets in {wb.Name} ({sortMode})";
            }
            catch (Exception ex)
            {
                _lblStatus.Text = $"Sort error: {ex.Message}";
            }
        }

        // --- FEATURE 4: TAB COLOR PICKER ---
        private void ChangeSheetTabColor()
        {
            try
            {
                if (_lstSheets.SelectedIndex >= 0 && _lstSheets.SelectedIndex < _lstSheets.Items.Count)
                {
                    if (_lstSheets.Items[_lstSheets.SelectedIndex] is SheetItem targetSheet && targetSheet.SheetRef is Excel.Worksheet ws)
                    {
                        using (ColorDialog cd = new ColorDialog())
                        {
                            cd.AllowFullOpen = true;
                            cd.FullOpen = true;
                            try
                            {
                                cd.Color = ColorTranslator.FromHtml(targetSheet.TabColorHex);
                            }
                            catch { }

                            if (cd.ShowDialog() == DialogResult.OK)
                            {
                                ws.Tab.Color = ColorTranslator.ToOle(cd.Color);
                                RefreshData();
                                _lblStatus.Text = $"Changed tab color for sheet: {ws.Name}";
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _lblStatus.Text = $"Color change error: {ex.Message}";
            }
        }

        private void ResetSheetTabColor()
        {
            try
            {
                if (_lstSheets.SelectedIndex >= 0 && _lstSheets.SelectedIndex < _lstSheets.Items.Count)
                {
                    if (_lstSheets.Items[_lstSheets.SelectedIndex] is SheetItem targetSheet && targetSheet.SheetRef is Excel.Worksheet ws)
                    {
                        ws.Tab.ColorIndex = Excel.XlColorIndex.xlColorIndexNone;
                        RefreshData();
                        _lblStatus.Text = $"Reset tab color for sheet: {ws.Name}";
                    }
                }
            }
            catch (Exception ex)
            {
                _lblStatus.Text = $"Color reset error: {ex.Message}";
            }
        }

        // --- FEATURE 6: EXPORT TO PDF & NEW WORKBOOK ---
        private void ExportSheetToPdf()
        {
            try
            {
                if (_lstSheets.SelectedIndex >= 0 && _lstSheets.SelectedIndex < _lstSheets.Items.Count)
                {
                    if (_lstSheets.Items[_lstSheets.SelectedIndex] is SheetItem targetSheet && targetSheet.SheetRef is Excel.Worksheet ws)
                    {
                        using (SaveFileDialog sfd = new SaveFileDialog())
                        {
                            sfd.Filter = "PDF File (*.pdf)|*.pdf";
                            sfd.FileName = $"{ws.Name}.pdf";
                            if (sfd.ShowDialog() == DialogResult.OK)
                            {
                                ws.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, sfd.FileName);
                                _lblStatus.Text = $"Exported '{ws.Name}' to PDF: {Path.GetFileName(sfd.FileName)}";
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _lblStatus.Text = $"PDF export error: {ex.Message}";
            }
        }

        private void ExportSheetToNewWorkbook()
        {
            try
            {
                if (_lstSheets.SelectedIndex >= 0 && _lstSheets.SelectedIndex < _lstSheets.Items.Count)
                {
                    if (_lstSheets.Items[_lstSheets.SelectedIndex] is SheetItem targetSheet && targetSheet.SheetRef is Excel.Worksheet ws)
                    {
                        using (SaveFileDialog sfd = new SaveFileDialog())
                        {
                            sfd.Filter = "Excel Workbook (*.xlsx)|*.xlsx";
                            sfd.FileName = $"{ws.Name}.xlsx";
                            if (sfd.ShowDialog() == DialogResult.OK)
                            {
                                ws.Copy();
                                var excelApp = (Excel.Application)ExcelDnaUtil.Application;
                                if (excelApp.ActiveWorkbook != null)
                                {
                                    excelApp.ActiveWorkbook.SaveAs(sfd.FileName);
                                }
                                RefreshData();
                                _lblStatus.Text = $"Exported '{ws.Name}' to new file: {Path.GetFileName(sfd.FileName)}";
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _lblStatus.Text = $"Excel export error: {ex.Message}";
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
            Color itemText = isSelected ? Color.White : _textColor;

            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;

            using (var brush = new SolidBrush(itemBg))
            {
                e.Graphics.FillRectangle(brush, new Rectangle(e.Bounds.X + 2, e.Bounds.Y + 2, e.Bounds.Width - 4, e.Bounds.Height - 4));
            }

            TextRenderer.DrawText(e.Graphics, "📄", new Font("Segoe UI", 10.5f), new Point(e.Bounds.X + 6, e.Bounds.Y + 7), itemText);

            using (var fontName = new Font("Segoe UI", 9.5f, FontStyle.Bold))
            {
                Rectangle nameRect = new Rectangle(e.Bounds.X + 30, e.Bounds.Y + 4, e.Bounds.Width - 36, e.Bounds.Height - 8);
                TextRenderer.DrawText(e.Graphics, item.Name, fontName, nameRect, itemText, TextFormatFlags.VerticalCenter | TextFormatFlags.Left | TextFormatFlags.EndEllipsis);
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
                itemBg = Color.FromArgb(79, 70, 229);
                textCol = Color.White;
            }
            else if (item.HasCustomTabColor)
            {
                itemBg = tabColor;
                double brightness = (tabColor.R * 0.299 + tabColor.G * 0.587 + tabColor.B * 0.114);
                textCol = brightness > 160 ? Color.FromArgb(15, 23, 42) : Color.White;
            }
            else
            {
                itemBg = _cardColor;
                textCol = _textColor;
            }

            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;

            using (var brush = new SolidBrush(itemBg))
            {
                e.Graphics.FillRectangle(brush, new Rectangle(e.Bounds.X + 2, e.Bounds.Y + 2, e.Bounds.Width - 4, e.Bounds.Height - 4));
            }

            if (!item.HasCustomTabColor && !isSelected)
            {
                using (var tabBrush = new SolidBrush(tabColor))
                {
                    e.Graphics.FillRectangle(tabBrush, new Rectangle(e.Bounds.X + 6, e.Bounds.Y + 6, 5, e.Bounds.Height - 12));
                }
            }

            int badgeOffset = 10;
            if (!item.IsVisible) badgeOffset += 52;
            if (item.IsProtected) badgeOffset += 32;

            int leftOffset = (item.HasCustomTabColor || isSelected) ? 10 : 18;
            using (var fontName = new Font("Segoe UI", 9.5f, FontStyle.Bold))
            {
                Rectangle nameRect = new Rectangle(e.Bounds.X + leftOffset, e.Bounds.Y + 4, e.Bounds.Width - leftOffset - badgeOffset, e.Bounds.Height - 8);
                TextRenderer.DrawText(e.Graphics, item.Name, fontName, nameRect, textCol, TextFormatFlags.VerticalCenter | TextFormatFlags.Left | TextFormatFlags.EndEllipsis);
            }

            int currentRight = e.Bounds.Right - 8;

            if (!item.IsVisible)
            {
                using (var fontBadge = new Font("Segoe UI", 7.5f, FontStyle.Bold))
                using (var badgeBg = new SolidBrush(Color.FromArgb(239, 68, 68)))
                {
                    Rectangle badgeRect = new Rectangle(currentRight - 46, e.Bounds.Y + 8, 46, 18);
                    e.Graphics.FillRectangle(badgeBg, badgeRect);
                    TextRenderer.DrawText(e.Graphics, "Hidden", fontBadge, new Point(badgeRect.X + 5, badgeRect.Y + 2), Color.White);
                    currentRight -= 50;
                }
            }

            if (item.IsProtected)
            {
                using (var fontBadge = new Font("Segoe UI", 8f, FontStyle.Bold))
                using (var badgeBg = new SolidBrush(Color.FromArgb(245, 158, 11)))
                {
                    Rectangle badgeRect = new Rectangle(currentRight - 26, e.Bounds.Y + 8, 24, 18);
                    e.Graphics.FillRectangle(badgeBg, badgeRect);
                    TextRenderer.DrawText(e.Graphics, "🔒", fontBadge, new Point(badgeRect.X + 4, badgeRect.Y + 1), Color.White);
                }
            }
        }
    }
}
