using System;
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
        private static CustomTaskPane? _taskPane;
        private static TaskPaneHost? _taskPaneHost;
        private static Excel.Application? _excelApp;

        public void AutoOpen()
        {
            try
            {
                _excelApp = (Excel.Application)ExcelDnaUtil.Application;
                if (_excelApp != null)
                {
                    // Register Excel Application Event Listeners for Live Auto-Sync
                    _excelApp.WorkbookOpen += ExcelApp_WorkbookOpen;
                    _excelApp.WorkbookBeforeClose += ExcelApp_WorkbookBeforeClose;
                    _excelApp.WorkbookActivate += ExcelApp_WorkbookActivate;
                    _excelApp.SheetActivate += ExcelApp_SheetActivate;
                    _excelApp.WorkbookNewSheet += ExcelApp_WorkbookNewSheet;
                }

                // Try safe initialization after Excel finishes startup
                ExcelAsyncUtil.QueueAsMacro(() =>
                {
                    InitializeTaskPane();
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
                    _excelApp.SheetActivate -= ExcelApp_SheetActivate;
                    _excelApp.WorkbookNewSheet -= ExcelApp_WorkbookNewSheet;
                    _excelApp = null;
                }

                if (_taskPane != null)
                {
                    try { _taskPane.Visible = false; } catch { }
                    _taskPane = null;
                    _taskPaneHost = null;
                }
            }
            catch
            {
                // Cleanup silent
            }
        }

        public static void InitializeTaskPane()
        {
            try
            {
                if (_excelApp == null)
                {
                    _excelApp = (Excel.Application)ExcelDnaUtil.Application;
                }

                // If Excel has no workbooks open yet (e.g. Start Screen), defer creation until a workbook opens
                if (_excelApp == null || _excelApp.Workbooks == null || _excelApp.Workbooks.Count == 0)
                {
                    return;
                }

                // Check if existing TaskPane COM wrapper is dead/disconnected
                if (!IsTaskPaneValid())
                {
                    _taskPane = CustomTaskPaneFactory.CreateCustomTaskPane(typeof(TaskPaneHost), "Sheet & File Manager");
                    _taskPaneHost = _taskPane.ContentControl as TaskPaneHost;
                    _taskPane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight;
                    _taskPane.Width = 340;
                    _taskPane.Visible = true;
                }
                else if (_taskPane != null)
                {
                    _taskPane.Visible = true;
                }
            }
            catch (Exception ex)
            {
                _taskPane = null;
                _taskPaneHost = null;
                System.Diagnostics.Debug.WriteLine($"Could not create TaskPane: {ex.Message}");
            }
        }

        private static bool IsTaskPaneValid()
        {
            if (_taskPane == null) return false;
            try
            {
                // Touch property to verify COM object is alive
                var dummy = _taskPane.Visible;
                return true;
            }
            catch
            {
                _taskPane = null;
                _taskPaneHost = null;
                return false;
            }
        }

        public static void ToggleTaskPane()
        {
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                try
                {
                    if (!IsTaskPaneValid())
                    {
                        InitializeTaskPane();
                    }
                    else if (_taskPane != null)
                    {
                        _taskPane.Visible = !_taskPane.Visible;
                        if (_taskPane.Visible)
                        {
                            RefreshTaskPane();
                        }
                    }
                }
                catch
                {
                    _taskPane = null;
                    _taskPaneHost = null;
                    InitializeTaskPane();
                }
            });
        }

        public static void RefreshTaskPane()
        {
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                if (!IsTaskPaneValid())
                {
                    InitializeTaskPane();
                }
                _taskPaneHost?.ViewModel?.RefreshAllData();
            });
        }

        private void ExcelApp_WorkbookOpen(Excel.Workbook Wb) => TriggerRefresh();
        private void ExcelApp_WorkbookBeforeClose(Excel.Workbook Wb, ref bool Cancel) => TriggerRefresh();
        private void ExcelApp_WorkbookActivate(Excel.Workbook Wb) => TriggerRefresh();
        private void ExcelApp_SheetActivate(object Sh) => TriggerRefresh();
        private void ExcelApp_WorkbookNewSheet(Excel.Workbook Wb, object Sh) => TriggerRefresh();

        private static void TriggerRefresh()
        {
            RefreshTaskPane();
        }
    }
}
