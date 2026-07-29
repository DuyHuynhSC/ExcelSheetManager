using System;
using System.Runtime.InteropServices;
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

                InitializeTaskPane();
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
                    _taskPane.Visible = false;
                    _taskPane = null;
                }
            }
            catch
            {
                // Cleanup silent
            }
        }

        public static void InitializeTaskPane()
        {
            if (_taskPane == null)
            {
                _taskPaneHost = new TaskPaneHost();
                _taskPane = CustomTaskPaneFactory.CreateCustomTaskPane(_taskPaneHost, "Sheet & File Manager");
                _taskPane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight;
                _taskPane.Width = 340;
                _taskPane.Visible = true;
            }
            else
            {
                _taskPane.Visible = true;
            }
        }

        public static void ToggleTaskPane()
        {
            if (_taskPane == null)
            {
                InitializeTaskPane();
            }
            else
            {
                _taskPane.Visible = !_taskPane.Visible;
            }
        }

        public static void RefreshTaskPane()
        {
            _taskPaneHost?.ViewModel?.RefreshAllData();
        }

        private void ExcelApp_WorkbookOpen(Excel.Workbook Wb) => TriggerRefresh();
        private void ExcelApp_WorkbookBeforeClose(Excel.Workbook Wb, ref bool Cancel) => TriggerRefresh();
        private void ExcelApp_WorkbookActivate(Excel.Workbook Wb) => TriggerRefresh();
        private void ExcelApp_SheetActivate(object Sh) => TriggerRefresh();
        private void ExcelApp_WorkbookNewSheet(Excel.Workbook Wb, object Sh) => TriggerRefresh();

        private static void TriggerRefresh()
        {
            // Run refresh safely via ExcelAsyncUtil
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                _taskPaneHost?.ViewModel?.RefreshAllData();
            });
        }
    }
}
