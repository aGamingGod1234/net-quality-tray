using Microsoft.Win32;
using System.Drawing;

namespace NetQualitySentinel
{
    internal sealed class UiTheme
    {
        public bool IsDark { get; private set; }
        public Color AppBackground { get; private set; }
        public Color CardBackground { get; private set; }
        public Color CardBorder { get; private set; }
        public Color TextPrimary { get; private set; }
        public Color TextSecondary { get; private set; }
        public Color InputBackground { get; private set; }
        public Color InputBorder { get; private set; }
        public Color SecondaryButtonBackground { get; private set; }
        public Color SecondaryButtonBorder { get; private set; }
        public Color SecondaryButtonText { get; private set; }
        public Color PrimaryButtonBackground { get; private set; }
        public Color PrimaryButtonBorder { get; private set; }
        public Color PrimaryButtonText { get; private set; }
        public Color ChipSurface { get; private set; }
        public Color Accent { get; private set; }

        public static UiTheme GetCurrent()
        {
            bool isDark = IsDarkModeEnabled();
            return isDark ? CreateDark() : CreateLight();
        }

        private static bool IsDarkModeEnabled()
        {
            try
            {
                using (RegistryKey key = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize", false))
                {
                    if (key == null)
                    {
                        return false;
                    }

                    object value = key.GetValue("AppsUseLightTheme");
                    if (value is int)
                    {
                        return ((int)value) == 0;
                    }
                }
            }
            catch
            {
            }

            return false;
        }

        private static UiTheme CreateLight()
        {
            return new UiTheme
            {
                IsDark = false,
                AppBackground = Color.FromArgb(243, 246, 250),
                CardBackground = Color.White,
                CardBorder = Color.FromArgb(210, 218, 226),
                TextPrimary = Color.FromArgb(28, 33, 39),
                TextSecondary = Color.FromArgb(92, 102, 114),
                InputBackground = Color.White,
                InputBorder = Color.FromArgb(203, 211, 220),
                SecondaryButtonBackground = Color.FromArgb(247, 249, 252),
                SecondaryButtonBorder = Color.FromArgb(203, 211, 220),
                SecondaryButtonText = Color.FromArgb(28, 33, 39),
                PrimaryButtonBackground = Color.FromArgb(0, 95, 184),
                PrimaryButtonBorder = Color.FromArgb(0, 86, 170),
                PrimaryButtonText = Color.White,
                ChipSurface = Color.FromArgb(247, 249, 252),
                Accent = Color.FromArgb(0, 95, 184)
            };
        }

        private static UiTheme CreateDark()
        {
            return new UiTheme
            {
                IsDark = true,
                AppBackground = Color.FromArgb(32, 36, 43),
                CardBackground = Color.FromArgb(43, 48, 58),
                CardBorder = Color.FromArgb(76, 84, 97),
                TextPrimary = Color.FromArgb(236, 240, 246),
                TextSecondary = Color.FromArgb(173, 183, 196),
                InputBackground = Color.FromArgb(37, 42, 50),
                InputBorder = Color.FromArgb(88, 96, 111),
                SecondaryButtonBackground = Color.FromArgb(56, 63, 75),
                SecondaryButtonBorder = Color.FromArgb(86, 95, 109),
                SecondaryButtonText = Color.FromArgb(236, 240, 246),
                PrimaryButtonBackground = Color.FromArgb(78, 180, 255),
                PrimaryButtonBorder = Color.FromArgb(67, 164, 236),
                PrimaryButtonText = Color.FromArgb(20, 27, 35),
                ChipSurface = Color.FromArgb(37, 42, 50),
                Accent = Color.FromArgb(78, 180, 255)
            };
        }
    }
}
