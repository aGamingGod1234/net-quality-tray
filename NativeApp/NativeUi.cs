using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace NetQualitySentinel
{
    internal static class NativeUi
    {
        private const int DwmUseImmersiveDarkMode = 20;
        private const int DwmUseImmersiveDarkModeLegacy = 19;
        private static readonly IntPtr DpiAwareV2Context = new IntPtr(-4);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool SetProcessDPIAware();

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool SetProcessDpiAwarenessContext(IntPtr dpiContext);

        [DllImport("shcore.dll", SetLastError = true)]
        private static extern int SetProcessDpiAwareness(ProcessDpiAwareness awareness);

        [DllImport("dwmapi.dll")]
        private static extern int DwmSetWindowAttribute(IntPtr hwnd, int attribute, ref int value, int valueSize);

        internal static void EnableDpiAwareness()
        {
            try
            {
                if (SetProcessDpiAwarenessContext(DpiAwareV2Context))
                {
                    return;
                }
            }
            catch
            {
            }

            try
            {
                if (SetProcessDpiAwareness(ProcessDpiAwareness.ProcessPerMonitorDpiAware) == 0)
                {
                    return;
                }
            }
            catch
            {
            }

            try
            {
                SetProcessDPIAware();
            }
            catch
            {
            }
        }

        internal static void ApplyWindowTheme(Form form, bool dark)
        {
            if (form == null || !form.IsHandleCreated)
            {
                return;
            }

            int useDark = dark ? 1 : 0;
            try
            {
                DwmSetWindowAttribute(form.Handle, DwmUseImmersiveDarkMode, ref useDark, sizeof(int));
            }
            catch
            {
            }

            try
            {
                DwmSetWindowAttribute(form.Handle, DwmUseImmersiveDarkModeLegacy, ref useDark, sizeof(int));
            }
            catch
            {
            }
        }

        private enum ProcessDpiAwareness
        {
            ProcessDpiUnaware = 0,
            ProcessSystemDpiAware = 1,
            ProcessPerMonitorDpiAware = 2
        }
    }
}
