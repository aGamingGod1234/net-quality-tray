using Microsoft.Win32;

namespace NetQualitySentinel
{
    internal static class StartupRegistration
    {
        private const string RunKeyPath = @"Software\Microsoft\Windows\CurrentVersion\Run";

        public static bool IsEnabled(string valueName)
        {
            try
            {
                using (RegistryKey key = Registry.CurrentUser.OpenSubKey(RunKeyPath, false))
                {
                    if (key == null)
                    {
                        return false;
                    }

                    object value = key.GetValue(valueName);
                    return value is string && !string.IsNullOrWhiteSpace((string)value);
                }
            }
            catch
            {
                return false;
            }
        }

        public static void SetEnabled(string valueName, string exePath, bool enabled)
        {
            using (RegistryKey key = Registry.CurrentUser.CreateSubKey(RunKeyPath))
            {
                if (key == null)
                {
                    return;
                }

                if (enabled)
                {
                    string launch = "\"" + exePath + "\"";
                    key.SetValue(valueName, launch, RegistryValueKind.String);
                }
                else
                {
                    key.DeleteValue(valueName, false);
                }
            }
        }
    }
}
