using System;
using System.IO;
using Microsoft.Win32;

namespace ExcelSheetManager.Helpers
{
    public static class SettingsHelper
    {
        private const string KeyPath = @"SOFTWARE\ExcelSheetManager";
        private const string ValueName = "IsTaskPaneVisible";

        private static string FallbackFilePath
        {
            get
            {
                string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                string dir = Path.Combine(appData, "ExcelSheetManager");
                if (!Directory.Exists(dir))
                {
                    Directory.CreateDirectory(dir);
                }
                return Path.Combine(dir, "taskpane_state.txt");
            }
        }

        public static bool GetIsTaskPaneVisible(bool defaultValue = true)
        {
            try
            {
                using var key = Registry.CurrentUser.OpenSubKey(KeyPath);
                if (key != null)
                {
                    object val = key.GetValue(ValueName, defaultValue ? 1 : 0);
                    return Convert.ToInt32(val) != 0;
                }
            }
            catch { }

            // Fallback to text file in %AppData%
            try
            {
                if (File.Exists(FallbackFilePath))
                {
                    string text = File.ReadAllText(FallbackFilePath).Trim();
                    if (bool.TryParse(text, out bool result))
                    {
                        return result;
                    }
                }
            }
            catch { }

            return defaultValue;
        }

        public static void SetIsTaskPaneVisible(bool isVisible)
        {
            // 1. Save to Registry
            try
            {
                using var key = Registry.CurrentUser.CreateSubKey(KeyPath);
                key?.SetValue(ValueName, isVisible ? 1 : 0, RegistryValueKind.DWord);
            }
            catch { }

            // 2. Save to Fallback Text File
            try
            {
                File.WriteAllText(FallbackFilePath, isVisible.ToString());
            }
            catch { }
        }
    }
}
