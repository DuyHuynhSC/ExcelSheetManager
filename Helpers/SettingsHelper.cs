using System;
using System.IO;
using Microsoft.Win32;

namespace ExcelSheetManager.Helpers
{
    internal static class SettingsHelper
    {
        private const string KeyPath = @"SOFTWARE\ExcelSheetManager";
        private const string VisibilityValueName = "IsTaskPaneVisible";
        private const string ThemeValueName = "IsDarkTheme";
        private const string AiBaseUrlValueName = "AiBaseUrl";
        private const string AiApiKeyValueName = "AiApiKey";
        private const string AiModelNameValueName = "AiModelName";

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
        private static string AiSettingsFilePath => Path.Combine(AppDataDir, "ai_settings.txt");

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

        // --- LOCAL OPENAI AI SETTINGS ---
        public static string GetAiBaseUrl(string defaultValue = "http://localhost:11434/v1")
        {
            try
            {
                using var key = Registry.CurrentUser.OpenSubKey(KeyPath);
                if (key != null)
                {
                    object val = key.GetValue(AiBaseUrlValueName, defaultValue);
                    if (val != null) return val.ToString();
                }
            }
            catch { }

            return defaultValue;
        }

        public static void SetAiBaseUrl(string baseUrl)
        {
            try
            {
                using var key = Registry.CurrentUser.CreateSubKey(KeyPath);
                key?.SetValue(AiBaseUrlValueName, baseUrl, RegistryValueKind.String);
            }
            catch { }
        }

        public static string GetAiApiKey(string defaultValue = "local")
        {
            try
            {
                using var key = Registry.CurrentUser.OpenSubKey(KeyPath);
                if (key != null)
                {
                    object val = key.GetValue(AiApiKeyValueName, defaultValue);
                    if (val != null) return val.ToString();
                }
            }
            catch { }

            return defaultValue;
        }

        public static void SetAiApiKey(string apiKey)
        {
            try
            {
                using var key = Registry.CurrentUser.CreateSubKey(KeyPath);
                key?.SetValue(AiApiKeyValueName, apiKey, RegistryValueKind.String);
            }
            catch { }
        }

        public static string GetAiModelName(string defaultValue = "llama3")
        {
            try
            {
                using var key = Registry.CurrentUser.OpenSubKey(KeyPath);
                if (key != null)
                {
                    object val = key.GetValue(AiModelNameValueName, defaultValue);
                    if (val != null) return val.ToString();
                }
            }
            catch { }

            return defaultValue;
        }

        public static void SetAiModelName(string modelName)
        {
            try
            {
                using var key = Registry.CurrentUser.CreateSubKey(KeyPath);
                key?.SetValue(AiModelNameValueName, modelName, RegistryValueKind.String);
            }
            catch { }
        }

        // --- GLOSSARY (JP - VN) SETTINGS ---
        private const string GlossaryValueName = "AiGlossaryJpVn";
        private static string GlossaryFilePath => Path.Combine(AppDataDir, "glossary_jp_vn.txt");

        public static string GetGlossaryText()
        {
            string defaultGlossary = "売上=Doanh thu\r\n利益=Lợi nhuận\r\n勘定科目=Tài khoản kế toán\r\n残高=Số dư\r\n振込=Chuyển khoản\r\n税込=Đã bao gồm thuế\r\n税抜=Chưa bao gồm thuế\r\n請求書=Hóa đơn\r\n納品書=Biên bản giao hàng\r\n発注書=Đơn đặt hàng";

            try
            {
                using var key = Registry.CurrentUser.OpenSubKey(KeyPath);
                if (key != null)
                {
                    object val = key.GetValue(GlossaryValueName, null);
                    if (val != null && !string.IsNullOrEmpty(val.ToString())) return val.ToString()!;
                }
            }
            catch { }

            try
            {
                if (File.Exists(GlossaryFilePath))
                {
                    string text = File.ReadAllText(GlossaryFilePath).Trim();
                    if (!string.IsNullOrEmpty(text)) return text;
                }
            }
            catch { }

            return defaultGlossary;
        }

        public static void SetGlossaryText(string text)
        {
            try
            {
                using var key = Registry.CurrentUser.CreateSubKey(KeyPath);
                key?.SetValue(GlossaryValueName, text, RegistryValueKind.String);
            }
            catch { }

            try
            {
                File.WriteAllText(GlossaryFilePath, text);
            }
            catch { }
        }
    }
}
