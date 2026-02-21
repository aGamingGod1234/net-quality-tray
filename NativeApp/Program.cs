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

            bool createdNew = false;
            using (Mutex mutex = new Mutex(true, MutexName, out createdNew))
            {
                if (!createdNew)
                {
                    MessageBox.Show(
                        AppName + " is already running.",
                        AppName,
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information
                    );
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
                        if (string.Equals(args[i], "--open-graph", StringComparison.OrdinalIgnoreCase))
                        {
                            openGraphOnStart = true;
                            continue;
                        }
                        if (string.Equals(args[i], "--open-settings", StringComparison.OrdinalIgnoreCase))
                        {
                            openSettingsOnStart = true;
                        }
                    }
                }

                Application.Run(new SentinelAppContext(AppName, appDir, settingsPath, iconPath, openGraphOnStart, openSettingsOnStart));
            }
        }
    }
}
