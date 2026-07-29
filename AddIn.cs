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

                // Initialize taskpane for active window on startup
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
                _excelApp = null;
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
                    // Create CustomTaskPane using pure WinForms TaskPaneControl directly bound to targetWin
                    CustomTaskPane ctp = CustomTaskPaneFactory.CreateCustomTaskPane(typeof(TaskPaneControl), "Sheet & File Manager", targetWin);
                    
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
                        if (ctp.Visible && ctp.ContentControl is TaskPaneControl control)
                        {
                            ctp.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight;
                            ctp.Width = 340;
                            control.RefreshData();
                        }
                    }
                    else
                    {
                        var newCtp = EnsureTaskPaneForWindow(activeWin, createIfMissing: true);
                        if (newCtp?.ContentControl is TaskPaneControl control)
                        {
                            control.RefreshData();
                        }
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
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                if (_excelApp == null) _excelApp = (Excel.Application)ExcelDnaUtil.Application;
                Excel.Window? activeWin = _excelApp?.ActiveWindow;
                if (activeWin != null && _taskPaneMap.TryGetValue(activeWin.Hwnd, out var ctp) && IsTaskPaneAlive(ctp))
                {
                    if (ctp.ContentControl is TaskPaneControl control)
                    {
                        control.RefreshData();
                    }
                }
                else
                {
                    var newCtp = EnsureTaskPaneForWindow(activeWin, createIfMissing: true);
                    if (newCtp?.ContentControl is TaskPaneControl control)
                    {
                        control.RefreshData();
                    }
                }
            });
        }
    }
}
