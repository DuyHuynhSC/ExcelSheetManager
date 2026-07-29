using System;
using Microsoft.Win32;

namespace ExcelSheetManager.Helpers
{
    public static class SettingsHelper
    {
        private const string KeyPath = @"SOFTWARE\ExcelSheetManager";
        private const string ValueName = "IsTaskPaneVisible";

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
            return defaultValue;
        }

        public static void SetIsTaskPaneVisible(bool isVisible)
        {
            try
            {
                using var key = Registry.CurrentUser.CreateSubKey(KeyPath);
                key?.SetValue(ValueName, isVisible ? 1 : 0, RegistryValueKind.DWord);
            }
            catch { }
        }
    }
}
