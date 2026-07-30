using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using ExcelSheetManager.Helpers;
using ExcelSheetManager.Views;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelSheetManager
{
    public class AddIn : IExcelAddIn
    {
        private static readonly Dictionary<int, CustomTaskPane> _taskPaneMap = new Dictionary<int, CustomTaskPane>();
        private static Excel.Application? _excelApp;
        private static System.Windows.Forms.Timer? _startupTimer;

        public void AutoOpen()
        {
            try
            {
                _excelApp = (Excel.Application)ExcelDnaUtil.Application;
                if (_excelApp != null)
                {
                    // Register WorkbookOpen and WorkbookBeforeClose
                    _excelApp.WorkbookOpen += ExcelApp_WorkbookOpen;
                    _excelApp.WorkbookBeforeClose += ExcelApp_WorkbookBeforeClose;
                }

                // Run immediate attempt
                TryInitializeStartupTaskPane();

                // Start short polling timer (every 200ms up to 3 sec) until Excel's main window finishes initializing
                int attempts = 0;
                _startupTimer = new System.Windows.Forms.Timer
                {
                    Interval = 200
                };
                _startupTimer.Tick += (s, e) =>
                {
                    attempts++;
                    bool initialized = TryInitializeStartupTaskPane();
                    if (initialized || attempts > 15) // Max 3 seconds
                    {
                        if (_startupTimer != null)
                        {
                            _startupTimer.Stop();
                            _startupTimer.Dispose();
                            _startupTimer = null;
                        }
                    }
                };
                _startupTimer.Start();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"AutoOpen error: {ex.Message}");
            }
        }

        private static bool TryInitializeStartupTaskPane()
        {
            try
            {
                if (_excelApp == null) _excelApp = (Excel.Application)ExcelDnaUtil.Application;
                var activeWin = _excelApp?.ActiveWindow;
                if (activeWin != null)
                {
                    EnsureTaskPaneForWindow(activeWin, createIfMissing: true);
                    return true;
                }
            }
            catch { }
            return false;
        }

        public void AutoClose()
        {
            try
            {
                if (_startupTimer != null)
                {
                    _startupTimer.Stop();
                    _startupTimer.Dispose();
                    _startupTimer = null;
                }

                if (_excelApp != null)
                {
                    _excelApp.WorkbookOpen -= ExcelApp_WorkbookOpen;
                    _excelApp.WorkbookBeforeClose -= ExcelApp_WorkbookBeforeClose;
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

        private void ExcelApp_WorkbookBeforeClose(Excel.Workbook Wb, ref bool Cancel)
        {
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                try
                {
                    RefreshAllTaskPanes();
                }
                catch { }
            });
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

                string parentWbName = string.Empty;
                try
                {
                    if (targetWin.Parent is Excel.Workbook parentWb)
                    {
                        parentWbName = parentWb.Name;
                    }
                }
                catch { }

                if (_taskPaneMap.TryGetValue(hwnd, out var existingCtp))
                {
                    if (IsTaskPaneAlive(existingCtp))
                    {
                        if (existingCtp.ContentControl is TaskPaneControl existingControl && !string.IsNullOrEmpty(parentWbName))
                        {
                            existingControl.BoundWorkbookName = parentWbName;
                        }
                        return existingCtp;
                    }
                    else
                    {
                        _taskPaneMap.Remove(hwnd);
                    }
                }

                if (createIfMissing)
                {
                    bool shouldBeVisible = SettingsHelper.GetIsTaskPaneVisible(defaultValue: true);

                    // Create CustomTaskPane with Title "Navigation" bound to targetWin
                    CustomTaskPane ctp = CustomTaskPaneFactory.CreateCustomTaskPane(typeof(TaskPaneControl), "Navigation", targetWin);
                    
                    if (ctp.ContentControl is TaskPaneControl control && !string.IsNullOrEmpty(parentWbName))
                    {
                        control.BoundWorkbookName = parentWbName;
                    }

                    // Enforce Left side panel docking both before and after Visible
                    ctp.DockPosition = MsoCTPDockPosition.msoCTPDockPositionLeft;
                    ctp.Width = 340;
                    ctp.Visible = shouldBeVisible;

                    if (shouldBeVisible)
                    {
                        ctp.DockPosition = MsoCTPDockPosition.msoCTPDockPositionLeft;
                        ctp.Width = 340;
                    }

                    // Event listener to automatically persist visibility whenever user toggles or closes taskpane
                    ctp.VisibleStateChange += (senderPane) =>
                    {
                        try
                        {
                            if (senderPane is CustomTaskPane pane)
                            {
                                SettingsHelper.SetIsTaskPaneVisible(pane.Visible);
                            }
                        }
                        catch { }
                    };

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

        public static void SyncSelectedWorkbookInAllTaskPanes(string wbName)
        {
            if (string.IsNullOrEmpty(wbName)) return;

            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                foreach (var kvp in _taskPaneMap)
                {
                    try
                    {
                        if (IsTaskPaneAlive(kvp.Value) && kvp.Value.ContentControl is TaskPaneControl control)
                        {
                            control.SelectWorkbookByName(wbName);
                        }
                    }
                    catch { }
                }
            });
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
                        if (kvp.Value.ContentControl is TaskPaneControl control)
                        {
                            control.RefreshData();
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
                        SettingsHelper.SetIsTaskPaneVisible(ctp.Visible);

                        if (ctp.Visible && ctp.ContentControl is TaskPaneControl control)
                        {
                            try
                            {
                                if (activeWin.Parent is Excel.Workbook wb) control.BoundWorkbookName = wb.Name;
                            }
                            catch { }

                            ctp.DockPosition = MsoCTPDockPosition.msoCTPDockPositionLeft;
                            ctp.Width = 340;
                            control.RefreshData();
                        }
                    }
                    else
                    {
                        var newCtp = EnsureTaskPaneForWindow(activeWin, createIfMissing: true);
                        if (newCtp != null)
                        {
                            newCtp.Visible = true;
                            SettingsHelper.SetIsTaskPaneVisible(true);
                            if (newCtp.ContentControl is TaskPaneControl control)
                            {
                                control.RefreshData();
                            }
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
            RefreshAllTaskPanes();
        }

        [ExcelCommand(ShortCut = "^+F")]
        public static void ShowQuickJumpWindow()
        {
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                QuickJumpForm.ShowForm();
            });
        }

        [ExcelFunction(Description = "Ask Local OpenAI AI model a prompt with optional cell value", Category = "Excel Sheet Manager AI")]
        public static object ASK_AI(string prompt, object cellValue)
        {
            return ExcelAsyncUtil.Run("ASK_AI", new object[] { prompt, cellValue }, () =>
            {
                try
                {
                    string valStr = (cellValue != null && cellValue is not ExcelDna.Integration.ExcelEmpty) ? cellValue.ToString() : string.Empty;
                    string fullPrompt = string.IsNullOrEmpty(valStr) ? prompt : $"{prompt}: \"{valStr}\"";
                    var task = System.Threading.Tasks.Task.Run(() => AiService.GetCompletionAsync(fullPrompt, "You are a helpful AI assistant integrated inside Microsoft Excel. Respond concisely."));
                    task.Wait(15000);
                    return task.Result;
                }
                catch (Exception ex)
                {
                    return $"#AI_ERROR: {ex.Message}";
                }
            });
        }
    }
}
