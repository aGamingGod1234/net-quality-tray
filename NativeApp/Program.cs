using System;
using System.Threading;
using System.Windows.Forms;

namespace NetQualitySentinel
{
    internal static class Program
    {
        private const string AppName = "NQA";
        private const string MutexName = "Local\\NetQualitySentinel.NativeApp";

        [STAThread]
        private static void Main(string[] args)
        {
            NativeUi.EnableDpiAwareness();

            bool launchedFromStartup = HasArg(args, "--autorun");

            bool createdNew = false;
            using (Mutex mutex = new Mutex(true, MutexName, out createdNew))
            {
                if (!createdNew)
                {
                    if (!launchedFromStartup)
                    {
                        MessageBox.Show(
                            AppName + " is already running.",
                            AppName,
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information
                        );
                    }
                    return;
                }

                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);

                string appDir = AppDomain.CurrentDomain.BaseDirectory;
                string settingsPath = System.IO.Path.Combine(appDir, "settings.json");
                string iconPath = System.IO.Path.Combine(appDir, "assets", "NetQualitySentinel.ico");
                bool openGraphOnStart = false;
                bool openSettingsOnStart = false;
                if (args != null)
                {
                    for (int i = 0; i < args.Length; i++)
                    {
                        if (HasArg(args[i], "--open-graph"))
                        {
                            openGraphOnStart = true;
                            continue;
                        }
                        if (HasArg(args[i], "--open-settings"))
                        {
                            openSettingsOnStart = true;
                        }
                    }
                }

                Application.Run(new SentinelAppContext(AppName, appDir, settingsPath, iconPath, openGraphOnStart, openSettingsOnStart));
            }
        }

        private static bool HasArg(string[] args, string expected)
        {
            if (args == null)
            {
                return false;
            }

            for (int i = 0; i < args.Length; i++)
            {
                if (HasArg(args[i], expected))
                {
                    return true;
                }
            }

            return false;
        }

        private static bool HasArg(string value, string expected)
        {
            return string.Equals(value, expected, StringComparison.OrdinalIgnoreCase);
        }
    }
}
