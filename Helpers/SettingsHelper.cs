using System;
using System.IO;
using Microsoft.Win32;

namespace ExcelSheetManager.Helpers
{
    public static class SettingsHelper
    {
        private const string KeyPath = @"SOFTWARE\ExcelSheetManager";
        private const string VisibilityValueName = "IsTaskPaneVisible";
        private const string ThemeValueName = "IsDarkTheme";

        private static string AppDataDir
        {
            get
            {
                string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                string dir = Path.Combine(appData, "ExcelSheetManager");
                if (!Directory.Exists(dir))
                {
                    Directory.CreateDirectory(dir);
                }
                return dir;
            }
        }

        private static string VisibilityFilePath => Path.Combine(AppDataDir, "taskpane_state.txt");
        private static string ThemeFilePath => Path.Combine(AppDataDir, "theme_state.txt");

        public static bool GetIsTaskPaneVisible(bool defaultValue = true)
        {
            try
            {
                using var key = Registry.CurrentUser.OpenSubKey(KeyPath);
                if (key != null)
                {
                    object val = key.GetValue(VisibilityValueName, defaultValue ? 1 : 0);
                    return Convert.ToInt32(val) != 0;
                }
            }
            catch { }

            try
            {
                if (File.Exists(VisibilityFilePath))
                {
                    string text = File.ReadAllText(VisibilityFilePath).Trim();
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
            try
            {
                using var key = Registry.CurrentUser.CreateSubKey(KeyPath);
                key?.SetValue(VisibilityValueName, isVisible ? 1 : 0, RegistryValueKind.DWord);
            }
            catch { }

            try
            {
                File.WriteAllText(VisibilityFilePath, isVisible.ToString());
            }
            catch { }
        }

        public static bool GetIsDarkTheme(bool defaultValue = true)
        {
            try
            {
                using var key = Registry.CurrentUser.OpenSubKey(KeyPath);
                if (key != null)
                {
                    object val = key.GetValue(ThemeValueName, defaultValue ? 1 : 0);
                    return Convert.ToInt32(val) != 0;
                }
            }
            catch { }

            try
            {
                if (File.Exists(ThemeFilePath))
                {
                    string text = File.ReadAllText(ThemeFilePath).Trim();
                    if (bool.TryParse(text, out bool result))
                    {
                        return result;
                    }
                }
            }
            catch { }

            return defaultValue;
        }

        public static void SetIsDarkTheme(bool isDark)
        {
            try
            {
                using var key = Registry.CurrentUser.CreateSubKey(KeyPath);
                key?.SetValue(ThemeValueName, isDark ? 1 : 0, RegistryValueKind.DWord);
            }
            catch { }

            try
            {
                File.WriteAllText(ThemeFilePath, isDark.ToString());
            }
            catch { }
        }
    }
}
