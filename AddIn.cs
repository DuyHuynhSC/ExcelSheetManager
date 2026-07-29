using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using ExcelSheetManager.Views;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelSheetManager
{
    public class AddIn : IExcelAddIn
    {
        private static readonly Dictionary<int, CustomTaskPane> _taskPaneMap = new Dictionary<int, CustomTaskPane>();
        private static Excel.Application? _excelApp;

        public void AutoOpen()
        {
            try
            {
                _excelApp = (Excel.Application)ExcelDnaUtil.Application;
                if (_excelApp != null)
                {
                    // Register Excel Application Event Listeners for Live Auto-Sync across windows
                    _excelApp.WorkbookOpen += ExcelApp_WorkbookOpen;
                    _excelApp.WorkbookBeforeClose += ExcelApp_WorkbookBeforeClose;
                    _excelApp.WorkbookActivate += ExcelApp_WorkbookActivate;
                    _excelApp.WindowActivate += ExcelApp_WindowActivate;
                    _excelApp.SheetActivate += ExcelApp_SheetActivate;
                    _excelApp.WorkbookNewSheet += ExcelApp_WorkbookNewSheet;
                }

                // Initialize taskpane for active window
                ExcelAsyncUtil.QueueAsMacro(() =>
                {
                    EnsureTaskPaneForWindow(_excelApp?.ActiveWindow, createIfMissing: true);
                });
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"AutoOpen error: {ex.Message}");
            }
        }

        public void AutoClose()
        {
            try
            {
                if (_excelApp != null)
                {
                    _excelApp.WorkbookOpen -= ExcelApp_WorkbookOpen;
                    _excelApp.WorkbookBeforeClose -= ExcelApp_WorkbookBeforeClose;
                    _excelApp.WorkbookActivate -= ExcelApp_WorkbookActivate;
                    _excelApp.WindowActivate -= ExcelApp_WindowActivate;
                    _excelApp.SheetActivate -= ExcelApp_SheetActivate;
                    _excelApp.WorkbookNewSheet -= ExcelApp_WorkbookNewSheet;
                    _excelApp = null;
                }

                foreach (var ctp in _taskPaneMap.Values)
                {
                    try { ctp.Visible = false; ctp.Delete(); } catch { }
                }
                _taskPaneMap.Clear();
            }
            catch
            {
                // Cleanup silent
            }
        }

        public static CustomTaskPane? EnsureTaskPaneForWindow(Excel.Window? targetWin, bool createIfMissing = true)
        {
            try
            {
                if (_excelApp == null)
                {
                    _excelApp = (Excel.Application)ExcelDnaUtil.Application;
                }

                if (_excelApp == null || _excelApp.Workbooks == null || _excelApp.Workbooks.Count == 0)
                {
                    return null;
                }

                targetWin ??= _excelApp.ActiveWindow;
                if (targetWin == null) return null;

                int hwnd;
                try
                {
                    hwnd = targetWin.Hwnd;
                }
                catch
                {
                    return null;
                }

                if (_taskPaneMap.TryGetValue(hwnd, out var existingCtp))
                {
                    if (IsTaskPaneAlive(existingCtp))
                    {
                        return existingCtp;
                    }
                    else
                    {
                        _taskPaneMap.Remove(hwnd);
                    }
                }

                if (createIfMissing)
                {
                    // Create CustomTaskPane specifically bound to targetWin
                    CustomTaskPane ctp = CustomTaskPaneFactory.CreateCustomTaskPane(typeof(TaskPaneHost), "Sheet & File Manager", targetWin);
                    
                    // Enforce Right side panel docking both before and after Visible
                    ctp.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight;
                    ctp.Width = 340;
                    ctp.Visible = true;
                    ctp.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight;
                    ctp.Width = 340;

                    _taskPaneMap[hwnd] = ctp;
                    return ctp;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Taskpane Creation Exception: {ex.Message}");
            }
            return null;
        }

        private static bool IsTaskPaneAlive(CustomTaskPane ctp)
        {
            if (ctp == null) return false;
            try
            {
                var dummy = ctp.Visible;
                return true;
            }
            catch
            {
                return false;
            }
        }

        public static void ToggleTaskPane()
        {
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                try
                {
                    if (_excelApp == null) _excelApp = (Excel.Application)ExcelDnaUtil.Application;
                    Excel.Window? activeWin = _excelApp?.ActiveWindow;
                    if (activeWin == null) return;

                    int hwnd = activeWin.Hwnd;
                    if (_taskPaneMap.TryGetValue(hwnd, out var ctp) && IsTaskPaneAlive(ctp))
                    {
                        ctp.Visible = !ctp.Visible;
                        if (ctp.Visible)
                        {
                            ctp.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight;
                            ctp.Width = 340;
                            RefreshAllTaskPanes();
                        }
                    }
                    else
                    {
                        EnsureTaskPaneForWindow(activeWin, createIfMissing: true);
                        RefreshAllTaskPanes();
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Toggle error: {ex.Message}");
                }
            });
        }

        public static void RefreshTaskPane()
        {
            RefreshAllTaskPanes();
        }

        public static void RefreshAllTaskPanes()
        {
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                List<int> deadHwnds = new List<int>();
                foreach (var kvp in _taskPaneMap)
                {
                    if (IsTaskPaneAlive(kvp.Value))
                    {
                        if (kvp.Value.ContentControl is TaskPaneHost host && host.ViewModel != null)
                        {
                            host.ViewModel.RefreshAllData();
                        }
                    }
                    else
                    {
                        deadHwnds.Add(kvp.Key);
                    }
                }

                foreach (var h in deadHwnds)
                {
                    _taskPaneMap.Remove(h);
                }
            });
        }

        private void ExcelApp_WorkbookOpen(Excel.Workbook Wb)
        {
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                try
                {
                    Excel.Window? win = (Wb.Windows != null && Wb.Windows.Count > 0) ? (Excel.Window)Wb.Windows[1] : _excelApp?.ActiveWindow;
                    EnsureTaskPaneForWindow(win, createIfMissing: true);
                }
                catch { }
                RefreshAllTaskPanes();
            });
        }

        private void ExcelApp_WorkbookActivate(Excel.Workbook Wb)
        {
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                try
                {
                    Excel.Window? win = (Wb.Windows != null && Wb.Windows.Count > 0) ? (Excel.Window)Wb.Windows[1] : _excelApp?.ActiveWindow;
                    EnsureTaskPaneForWindow(win, createIfMissing: true);
                }
                catch { }
                RefreshAllTaskPanes();
            });
        }

        private void ExcelApp_WindowActivate(Excel.Workbook Wb, Excel.Window Wn)
        {
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                EnsureTaskPaneForWindow(Wn, createIfMissing: true);
                RefreshAllTaskPanes();
            });
        }

        private void ExcelApp_WorkbookBeforeClose(Excel.Workbook Wb, ref bool Cancel)
        {
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                RefreshAllTaskPanes();
            });
        }

        private void ExcelApp_SheetActivate(object Sh)
        {
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                RefreshAllTaskPanes();
            });
        }

        private void ExcelApp_WorkbookNewSheet(Excel.Workbook Wb, object Sh)
        {
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                RefreshAllTaskPanes();
            });
        }
    }
}
