using System;
using System.IO;
using System.Runtime.InteropServices;

namespace NetQualitySentinel
{
    internal static class StartMenuRegistration
    {
        private const string ShortcutName = "NQA.lnk";

        public static void EnsureShortcut(string appName, string exePath, string iconPath)
        {
            if (string.IsNullOrWhiteSpace(exePath) || !File.Exists(exePath))
            {
                return;
            }

            Type shellType = Type.GetTypeFromProgID("WScript.Shell");
            if (shellType == null)
            {
                return;
            }

            string[] programsDirs = new[]
            {
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.StartMenu), "Programs"),
                Environment.GetFolderPath(Environment.SpecialFolder.CommonPrograms)
            };

            object shell = null;
            try
            {
                shell = Activator.CreateInstance(shellType);
                for (int i = 0; i < programsDirs.Length; i++)
                {
                    string programsDir = programsDirs[i];
                    if (string.IsNullOrWhiteSpace(programsDir))
                    {
                        continue;
                    }

                    try
                    {
                        Directory.CreateDirectory(programsDir);
                        string shortcutPath = Path.Combine(programsDir, ShortcutName);
                        object shortcut = shellType.InvokeMember("CreateShortcut", System.Reflection.BindingFlags.InvokeMethod, null, shell, new object[] { shortcutPath });
                        if (shortcut == null)
                        {
                            continue;
                        }

                        try
                        {
                            Type shortcutType = shortcut.GetType();
                            shortcutType.InvokeMember("TargetPath", System.Reflection.BindingFlags.SetProperty, null, shortcut, new object[] { exePath });
                            shortcutType.InvokeMember("WorkingDirectory", System.Reflection.BindingFlags.SetProperty, null, shortcut, new object[] { Path.GetDirectoryName(exePath) ?? string.Empty });
                            shortcutType.InvokeMember("Description", System.Reflection.BindingFlags.SetProperty, null, shortcut, new object[] { appName + " Network Quality Monitor" });
                            if (!string.IsNullOrWhiteSpace(iconPath) && File.Exists(iconPath))
                            {
                                shortcutType.InvokeMember("IconLocation", System.Reflection.BindingFlags.SetProperty, null, shortcut, new object[] { iconPath });
                            }
                            shortcutType.InvokeMember("Save", System.Reflection.BindingFlags.InvokeMethod, null, shortcut, null);
                        }
                        finally
                        {
                            if (Marshal.IsComObject(shortcut))
                            {
                                Marshal.FinalReleaseComObject(shortcut);
                            }
                        }
                    }
                    catch
                    {
                    }
                }
            }
            catch
            {
            }
            finally
            {
                if (shell != null && Marshal.IsComObject(shell))
                {
                    Marshal.FinalReleaseComObject(shell);
                }
            }
        }
    }
}
