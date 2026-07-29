using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Input;
using ExcelDna.Integration;
using ExcelSheetManager.Helpers;
using ExcelSheetManager.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelSheetManager.ViewModels
{
    public class MainViewModel : ViewModelBase
    {
        private ObservableCollection<WorkbookItem> _allWorkbooks = new();
        private ObservableCollection<WorkbookItem> _filteredWorkbooks = new();
        private ObservableCollection<SheetItem> _allSheets = new();
        private ObservableCollection<SheetItem> _filteredSheets = new();

        private WorkbookItem? _selectedWorkbook;
        private SheetItem? _selectedSheet;
        private string _workbookFilterText = string.Empty;
        private string _sheetFilterText = string.Empty;
        private string _statusMessage = "Ready";
        private int _totalWorkbooksCount;
        private int _totalSheetsCount;

        public MainViewModel()
        {
            RefreshCommand = new RelayCommand(RefreshAllData);
            ActivateWorkbookCommand = new RelayCommand(param => ActivateWorkbook(param as WorkbookItem));
            ActivateSheetCommand = new RelayCommand(param => ActivateSheet(param as SheetItem));

            // Initial load
            RefreshAllData();
        }

        public ObservableCollection<WorkbookItem> FilteredWorkbooks
        {
            get => _filteredWorkbooks;
            set => SetProperty(ref _filteredWorkbooks, value);
        }

        public ObservableCollection<SheetItem> FilteredSheets
        {
            get => _filteredSheets;
            set => SetProperty(ref _filteredSheets, value);
        }

        public WorkbookItem? SelectedWorkbook
        {
            get => _selectedWorkbook;
            set
            {
                if (SetProperty(ref _selectedWorkbook, value))
                {
                    LoadSheetsForSelectedWorkbook();
                }
            }
        }

        public SheetItem? SelectedSheet
        {
            get => _selectedSheet;
            set => SetProperty(ref _selectedSheet, value);
        }

        public string WorkbookFilterText
        {
            get => _workbookFilterText;
            set
            {
                if (SetProperty(ref _workbookFilterText, value))
                {
                    ApplyWorkbookFilter();
                }
            }
        }

        public string SheetFilterText
        {
            get => _sheetFilterText;
            set
            {
                if (SetProperty(ref _sheetFilterText, value))
                {
                    ApplySheetFilter();
                }
            }
        }

        public string StatusMessage
        {
            get => _statusMessage;
            set => SetProperty(ref _statusMessage, value);
        }

        public int TotalWorkbooksCount
        {
            get => _totalWorkbooksCount;
            set => SetProperty(ref _totalWorkbooksCount, value);
        }

        public int TotalSheetsCount
        {
            get => _totalSheetsCount;
            set => SetProperty(ref _totalSheetsCount, value);
        }

        public ICommand RefreshCommand { get; }
        public ICommand ActivateWorkbookCommand { get; }
        public ICommand ActivateSheetCommand { get; }

        public void SelectWorkbookByName(string wbName)
        {
            if (string.IsNullOrEmpty(wbName)) return;

            var match = _allWorkbooks.FirstOrDefault(w => w.Name.Equals(wbName, StringComparison.OrdinalIgnoreCase));
            if (match != null && SelectedWorkbook?.Name != match.Name)
            {
                foreach (var w in _allWorkbooks)
                {
                    w.IsActive = (w.Name == match.Name);
                }
                OnPropertyChanged(nameof(FilteredWorkbooks));
                SelectedWorkbook = match;
            }
        }

        public void RefreshAllData()
        {
            try
            {
                var excelApp = (Excel.Application)ExcelDnaUtil.Application;
                if (excelApp == null)
                {
                    StatusMessage = "Cannot connect to Excel Application";
                    return;
                }

                string previousWorkbookName = SelectedWorkbook?.Name ?? string.Empty;
                var currentActiveWb = excelApp.ActiveWorkbook;

                _allWorkbooks.Clear();
                Excel.Workbooks workbooks = excelApp.Workbooks;

                foreach (Excel.Workbook wb in workbooks)
                {
                    try
                    {
                        bool isActive = currentActiveWb != null && wb.Name.Equals(currentActiveWb.Name, StringComparison.OrdinalIgnoreCase);
                        var item = new WorkbookItem
                        {
                            Name = wb.Name,
                            FullPath = wb.FullName,
                            WorkbookRef = wb,
                            SheetCount = wb.Sheets.Count,
                            IsActive = isActive
                        };
                        _allWorkbooks.Add(item);
                    }
                    catch
                    {
                        // Skip if workbook object cannot be read
                    }
                }

                TotalWorkbooksCount = _allWorkbooks.Count;
                ApplyWorkbookFilter();

                // Select previously selected workbook or default active workbook
                WorkbookItem? targetToSelect = _allWorkbooks.FirstOrDefault(w => w.Name.Equals(previousWorkbookName, StringComparison.OrdinalIgnoreCase))
                                              ?? _allWorkbooks.FirstOrDefault(w => w.IsActive)
                                              ?? _allWorkbooks.FirstOrDefault();

                SelectedWorkbook = targetToSelect;
                StatusMessage = $"Updated at {DateTime.Now:HH:mm:ss} ({_allWorkbooks.Count} files)";
            }
            catch (Exception ex)
            {
                StatusMessage = $"Refresh error: {ex.Message}";
            }
        }

        private void ApplyWorkbookFilter()
        {
            if (string.IsNullOrWhiteSpace(WorkbookFilterText))
            {
                FilteredWorkbooks = new ObservableCollection<WorkbookItem>(_allWorkbooks);
            }
            else
            {
                var filter = WorkbookFilterText.Trim();
                var filtered = _allWorkbooks.Where(w => w.Name.IndexOf(filter, StringComparison.OrdinalIgnoreCase) >= 0 ||
                                                        w.FullPath.IndexOf(filter, StringComparison.OrdinalIgnoreCase) >= 0);
                FilteredWorkbooks = new ObservableCollection<WorkbookItem>(filtered);
            }
        }

        private void LoadSheetsForSelectedWorkbook()
        {
            _allSheets.Clear();
            FilteredSheets.Clear();

            if (SelectedWorkbook == null || SelectedWorkbook.WorkbookRef is not Excel.Workbook wb)
            {
                TotalSheetsCount = 0;
                return;
            }

            try
            {
                Excel.Sheets sheets = wb.Sheets;
                object activeSheetObj = wb.ActiveSheet;
                string activeSheetName = string.Empty;

                if (activeSheetObj is Excel.Worksheet activeWs)
                {
                    activeSheetName = activeWs.Name;
                }

                foreach (object sheetObj in sheets)
                {
                    if (sheetObj is Excel.Worksheet ws)
                    {
                        var (colorHex, hasCustomColor) = ColorHelper.GetSheetTabColorHex(ws);
                        bool isActive = ws.Name.Equals(activeSheetName, StringComparison.OrdinalIgnoreCase);

                        var sheetItem = new SheetItem
                        {
                            Name = ws.Name,
                            ParentWorkbookName = wb.Name,
                            SheetRef = ws,
                            TabColorHex = colorHex,
                            HasCustomTabColor = hasCustomColor,
                            IsVisible = ws.Visible == Excel.XlSheetVisibility.xlSheetVisible,
                            IsActive = isActive,
                            SheetType = "Worksheet"
                        };
                        _allSheets.Add(sheetItem);
                    }
                    else if (sheetObj is Excel.Chart chart)
                    {
                        var sheetItem = new SheetItem
                        {
                            Name = chart.Name,
                            ParentWorkbookName = wb.Name,
                            SheetRef = chart,
                            TabColorHex = "#F59E0B",
                            HasCustomTabColor = true,
                            IsVisible = chart.Visible == Excel.XlSheetVisibility.xlSheetVisible,
                            IsActive = false,
                            SheetType = "Chart"
                        };
                        _allSheets.Add(sheetItem);
                    }
                }

                TotalSheetsCount = _allSheets.Count;
                ApplySheetFilter();
            }
            catch (Exception ex)
            {
                StatusMessage = $"Error loading sheets: {ex.Message}";
            }
        }

        private void ApplySheetFilter()
        {
            if (string.IsNullOrWhiteSpace(SheetFilterText))
            {
                FilteredSheets = new ObservableCollection<SheetItem>(_allSheets);
            }
            else
            {
                var filter = SheetFilterText.Trim();
                var filtered = _allSheets.Where(s => s.Name.IndexOf(filter, StringComparison.OrdinalIgnoreCase) >= 0);
                FilteredSheets = new ObservableCollection<SheetItem>(filtered);
            }
        }

        public void ActivateWorkbook(WorkbookItem? item)
        {
            if (item == null || item.WorkbookRef is not Excel.Workbook wb) return;

            try
            {
                wb.Activate();
                if (wb.Windows != null && wb.Windows.Count > 0)
                {
                    try
                    {
                        ((Excel.Window)wb.Windows[1]).Activate();
                    }
                    catch { }
                }

                foreach (var w in _allWorkbooks)
                {
                    w.IsActive = (w.Name == item.Name);
                }
                OnPropertyChanged(nameof(FilteredWorkbooks));
                StatusMessage = $"Activated file: {item.Name}";
            }
            catch (Exception ex)
            {
                StatusMessage = $"Could not activate file: {ex.Message}";
            }
        }

        public void ActivateSheet(SheetItem? item)
        {
            if (item == null) return;

            try
            {
                // First ensure workbook is active
                if (item.SheetRef is Excel.Worksheet ws)
                {
                    var parentWb = (Excel.Workbook)ws.Parent;
                    parentWb.Activate();
                    if (parentWb.Windows != null && parentWb.Windows.Count > 0)
                    {
                        try { ((Excel.Window)parentWb.Windows[1]).Activate(); } catch { }
                    }
                    ws.Activate();
                }
                else if (item.SheetRef is Excel.Chart chart)
                {
                    var parentWb = (Excel.Workbook)chart.Parent;
                    parentWb.Activate();
                    if (parentWb.Windows != null && parentWb.Windows.Count > 0)
                    {
                        try { ((Excel.Window)parentWb.Windows[1]).Activate(); } catch { }
                    }
                    chart.Activate();
                }

                // Update active state in UI
                foreach (var s in _allSheets)
                {
                    s.IsActive = (s.Name == item.Name);
                }
                OnPropertyChanged(nameof(FilteredSheets));

                StatusMessage = $"Focusing sheet: {item.Name}";
            }
            catch (Exception ex)
            {
                StatusMessage = $"Could not focus sheet: {ex.Message}";
            }
        }
    }
}
