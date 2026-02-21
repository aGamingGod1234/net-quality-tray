using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Text.RegularExpressions;
using System.Web.Script.Serialization;

namespace NetQualitySentinel
{
    internal sealed class AppConfig
    {
        public List<string> LatencyHosts { get; set; }
        public int LatencySamples { get; set; }
        public int LatencyTimeoutMs { get; set; }
        public int LatencyIntervalSec { get; set; }
        public List<string> DownloadEndpoints { get; set; }
        public List<string> UploadEndpoints { get; set; }
        public int DownloadSmallBytes { get; set; }
        public int DownloadFullBytes { get; set; }
        public int UploadSmallBytes { get; set; }
        public int UploadFullBytes { get; set; }
        public int DownloadProbeMaxMs { get; set; }
        public int UploadProbeMaxMs { get; set; }
        public int MaxEndpointBackoffSec { get; set; }
        public int HttpTimeoutSec { get; set; }
        public int SpeedIntervalGoodSec { get; set; }
        public int SpeedIntervalNormalSec { get; set; }
        public int SpeedIntervalPoorSec { get; set; }
        public int FullProbeIntervalSec { get; set; }
        public int HistorySize { get; set; }
        public int QualityHighMinScore { get; set; }
        public int QualityPoorMinScore { get; set; }
        public int QualityVeryPoorMinScore { get; set; }
        public string ColorHighHex { get; set; }
        public string ColorPoorHex { get; set; }
        public string ColorVeryPoorHex { get; set; }
        public string ColorBadHex { get; set; }
        public string ColorOfflineHex { get; set; }
        public string ColorPausedHex { get; set; }

        public AppConfig Clone()
        {
            return new AppConfig
            {
                LatencyHosts = new List<string>(LatencyHosts ?? new List<string>()),
                LatencySamples = LatencySamples,
                LatencyTimeoutMs = LatencyTimeoutMs,
                LatencyIntervalSec = LatencyIntervalSec,
                DownloadEndpoints = new List<string>(DownloadEndpoints ?? new List<string>()),
                UploadEndpoints = new List<string>(UploadEndpoints ?? new List<string>()),
                DownloadSmallBytes = DownloadSmallBytes,
                DownloadFullBytes = DownloadFullBytes,
                UploadSmallBytes = UploadSmallBytes,
                UploadFullBytes = UploadFullBytes,
                DownloadProbeMaxMs = DownloadProbeMaxMs,
                UploadProbeMaxMs = UploadProbeMaxMs,
                MaxEndpointBackoffSec = MaxEndpointBackoffSec,
                HttpTimeoutSec = HttpTimeoutSec,
                SpeedIntervalGoodSec = SpeedIntervalGoodSec,
                SpeedIntervalNormalSec = SpeedIntervalNormalSec,
                SpeedIntervalPoorSec = SpeedIntervalPoorSec,
                FullProbeIntervalSec = FullProbeIntervalSec,
                HistorySize = HistorySize,
                QualityHighMinScore = QualityHighMinScore,
                QualityPoorMinScore = QualityPoorMinScore,
                QualityVeryPoorMinScore = QualityVeryPoorMinScore,
                ColorHighHex = ColorHighHex,
                ColorPoorHex = ColorPoorHex,
                ColorVeryPoorHex = ColorVeryPoorHex,
                ColorBadHex = ColorBadHex,
                ColorOfflineHex = ColorOfflineHex,
                ColorPausedHex = ColorPausedHex
            };
        }

        public void Normalize()
        {
            LatencyHosts = NormalizeList(LatencyHosts, new List<string> { "1.1.1.1", "8.8.8.8", "9.9.9.9" });
            DownloadEndpoints = NormalizeDownloadEndpoints(NormalizeList(DownloadEndpoints, new List<string>
            {
                "https://speed.cloudflare.com/__down?bytes={bytes}",
                "https://speed.hetzner.de/100MB.bin",
                "https://proof.ovh.net/files/100Mb.dat"
            }));
            UploadEndpoints = NormalizeUploadEndpoints(NormalizeList(UploadEndpoints, new List<string>
            {
                "https://speed.cloudflare.com/__up",
                "https://postman-echo.com/post",
                "https://httpbin.org/post"
            }));

            LatencySamples = ClampInt(LatencySamples, 6, 20, 8);
            LatencyTimeoutMs = ClampInt(LatencyTimeoutMs, 300, 5000, 900);
            LatencyIntervalSec = ClampInt(LatencyIntervalSec, 1, 30, 4);

            DownloadSmallBytes = ClampInt(DownloadSmallBytes, 250000, 24000000, 4000000);
            DownloadFullBytes = ClampInt(DownloadFullBytes, 500000, 60000000, 18000000);
            UploadSmallBytes = ClampInt(UploadSmallBytes, 128000, 20000000, 3000000);
            UploadFullBytes = ClampInt(UploadFullBytes, 256000, 40000000, 12000000);
            DownloadProbeMaxMs = ClampInt(DownloadProbeMaxMs, 1200, 20000, 6000);
            UploadProbeMaxMs = ClampInt(UploadProbeMaxMs, 1200, 20000, 6000);
            MaxEndpointBackoffSec = ClampInt(MaxEndpointBackoffSec, 10, 900, 180);
            HttpTimeoutSec = ClampInt(HttpTimeoutSec, 4, 90, 15);

            SpeedIntervalPoorSec = ClampInt(SpeedIntervalPoorSec, 2, 120, 10);
            SpeedIntervalNormalSec = ClampInt(SpeedIntervalNormalSec, 3, 180, 18);
            SpeedIntervalGoodSec = ClampInt(SpeedIntervalGoodSec, 5, 300, 30);
            FullProbeIntervalSec = ClampInt(FullProbeIntervalSec, 30, 3600, 300);
            HistorySize = ClampInt(HistorySize, 6, 240, 20);

            if (DownloadFullBytes < DownloadSmallBytes)
            {
                DownloadFullBytes = DownloadSmallBytes;
            }
            if (UploadFullBytes < UploadSmallBytes)
            {
                UploadFullBytes = UploadSmallBytes;
            }
            if (SpeedIntervalNormalSec < SpeedIntervalPoorSec)
            {
                SpeedIntervalNormalSec = SpeedIntervalPoorSec;
            }
            if (SpeedIntervalGoodSec < SpeedIntervalNormalSec)
            {
                SpeedIntervalGoodSec = SpeedIntervalNormalSec;
            }

            QualityHighMinScore = ClampInt(QualityHighMinScore, 1, 100, 70);
            QualityPoorMinScore = ClampInt(QualityPoorMinScore, 0, 99, 45);
            QualityVeryPoorMinScore = ClampInt(QualityVeryPoorMinScore, 0, 98, 25);

            if (QualityPoorMinScore >= QualityHighMinScore)
            {
                QualityPoorMinScore = Math.Max(0, QualityHighMinScore - 1);
            }
            if (QualityVeryPoorMinScore >= QualityPoorMinScore)
            {
                QualityVeryPoorMinScore = Math.Max(0, QualityPoorMinScore - 1);
            }

            ColorHighHex = NormalizeHex(ColorHighHex, "#2ECC71");
            ColorPoorHex = NormalizeHex(ColorPoorHex, "#F1C40F");
            ColorVeryPoorHex = NormalizeHex(ColorVeryPoorHex, "#E67E22");
            ColorBadHex = NormalizeHex(ColorBadHex, "#E74C3C");
            ColorOfflineHex = NormalizeHex(ColorOfflineHex, "#951111");
            ColorPausedHex = NormalizeHex(ColorPausedHex, "#A0A0A0");
        }

        public Color GetTierColor(string tier)
        {
            string hex;
            switch ((tier ?? string.Empty).Trim())
            {
                case "High":
                    hex = ColorHighHex;
                    break;
                case "Poor":
                    hex = ColorPoorHex;
                    break;
                case "VeryPoor":
                    hex = ColorVeryPoorHex;
                    break;
                case "Bad":
                    hex = ColorBadHex;
                    break;
                case "Offline":
                    hex = ColorOfflineHex;
                    break;
                case "Paused":
                    hex = ColorPausedHex;
                    break;
                default:
                    hex = "#9AA6B2";
                    break;
            }

            try
            {
                return ColorTranslator.FromHtml(hex);
            }
            catch
            {
                return Color.Gray;
            }
        }

        public static AppConfig Load(string path)
        {
            AppConfig defaults = CreateDefaults();
            if (!File.Exists(path))
            {
                return defaults;
            }

            try
            {
                JavaScriptSerializer serializer = new JavaScriptSerializer();
                string json = File.ReadAllText(path);
                AppConfig parsed = serializer.Deserialize<AppConfig>(json);
                if (parsed == null)
                {
                    return defaults;
                }

                parsed.Normalize();
                return parsed;
            }
            catch
            {
                return defaults;
            }
        }

        public static void Save(string path, AppConfig config)
        {
            if (config == null)
            {
                return;
            }

            config.Normalize();
            string dir = Path.GetDirectoryName(path);
            if (!string.IsNullOrWhiteSpace(dir))
            {
                Directory.CreateDirectory(dir);
            }

            JavaScriptSerializer serializer = new JavaScriptSerializer();
            string json = serializer.Serialize(config);
            File.WriteAllText(path, json);
        }

        public static AppConfig CreateDefaults()
        {
            AppConfig config = new AppConfig();
            config.LatencyHosts = new List<string> { "1.1.1.1", "8.8.8.8", "9.9.9.9" };
            config.LatencySamples = 8;
            config.LatencyTimeoutMs = 900;
            config.LatencyIntervalSec = 4;
            config.DownloadEndpoints = new List<string>
            {
                "https://speed.cloudflare.com/__down?bytes={bytes}",
                "https://speed.hetzner.de/100MB.bin",
                "https://proof.ovh.net/files/100Mb.dat"
            };
            config.UploadEndpoints = new List<string>
            {
                "https://speed.cloudflare.com/__up",
                "https://postman-echo.com/post",
                "https://httpbin.org/post"
            };
            config.DownloadSmallBytes = 4000000;
            config.DownloadFullBytes = 18000000;
            config.UploadSmallBytes = 3000000;
            config.UploadFullBytes = 12000000;
            config.DownloadProbeMaxMs = 6000;
            config.UploadProbeMaxMs = 6000;
            config.MaxEndpointBackoffSec = 180;
            config.HttpTimeoutSec = 15;
            config.SpeedIntervalGoodSec = 30;
            config.SpeedIntervalNormalSec = 18;
            config.SpeedIntervalPoorSec = 10;
            config.FullProbeIntervalSec = 300;
            config.HistorySize = 20;
            config.QualityHighMinScore = 70;
            config.QualityPoorMinScore = 45;
            config.QualityVeryPoorMinScore = 25;
            config.ColorHighHex = "#2ECC71";
            config.ColorPoorHex = "#F1C40F";
            config.ColorVeryPoorHex = "#E67E22";
            config.ColorBadHex = "#E74C3C";
            config.ColorOfflineHex = "#951111";
            config.ColorPausedHex = "#A0A0A0";
            config.Normalize();
            return config;
        }

        private static string NormalizeHex(string hex, string fallback)
        {
            string raw = string.IsNullOrWhiteSpace(hex) ? fallback : hex.Trim();
            if (raw.StartsWith("#"))
            {
                raw = raw.Substring(1);
            }
            if (!Regex.IsMatch(raw, "^[0-9A-Fa-f]{6}$"))
            {
                raw = fallback.TrimStart('#');
            }
            return "#" + raw.ToUpperInvariant();
        }

        private static List<string> NormalizeList(List<string> values, List<string> fallback)
        {
            if (values == null)
            {
                return new List<string>(fallback);
            }

            List<string> normalized = new List<string>();
            foreach (string item in values)
            {
                if (string.IsNullOrWhiteSpace(item))
                {
                    continue;
                }

                string trimmed = item.Trim();
                if (!normalized.Contains(trimmed))
                {
                    normalized.Add(trimmed);
                }
            }

            if (normalized.Count == 0)
            {
                return new List<string>(fallback);
            }
            return normalized;
        }

        private static List<string> NormalizeDownloadEndpoints(List<string> endpoints)
        {
            List<string> normalized = new List<string>();
            List<string> source = endpoints ?? new List<string>();
            for (int i = 0; i < source.Count; i++)
            {
                string endpoint = source[i];
                if (string.IsNullOrWhiteSpace(endpoint))
                {
                    continue;
                }

                string mapped = endpoint.Trim();
                if (mapped.Equals("https://speed.hetzner.de/1MB.bin", StringComparison.OrdinalIgnoreCase))
                {
                    mapped = "https://speed.hetzner.de/100MB.bin";
                }
                else if (mapped.Equals("https://proof.ovh.net/files/1Mb.dat", StringComparison.OrdinalIgnoreCase))
                {
                    mapped = "https://proof.ovh.net/files/100Mb.dat";
                }

                if (!normalized.Contains(mapped))
                {
                    normalized.Add(mapped);
                }
            }

            EnsurePreferredFirst(normalized, "https://speed.cloudflare.com/__down?bytes={bytes}");
            return normalized;
        }

        private static List<string> NormalizeUploadEndpoints(List<string> endpoints)
        {
            List<string> normalized = new List<string>();
            List<string> source = endpoints ?? new List<string>();
            for (int i = 0; i < source.Count; i++)
            {
                string endpoint = source[i];
                if (string.IsNullOrWhiteSpace(endpoint))
                {
                    continue;
                }

                string mapped = endpoint.Trim();
                if (!normalized.Contains(mapped))
                {
                    normalized.Add(mapped);
                }
            }

            EnsurePreferredFirst(normalized, "https://speed.cloudflare.com/__up");
            return normalized;
        }

        private static void EnsurePreferredFirst(List<string> values, string preferred)
        {
            if (values == null || values.Count == 0 || string.IsNullOrWhiteSpace(preferred))
            {
                return;
            }

            string target = preferred.Trim();
            int index = -1;
            for (int i = 0; i < values.Count; i++)
            {
                if (string.Equals(values[i], target, StringComparison.OrdinalIgnoreCase))
                {
                    index = i;
                    break;
                }
            }

            if (index == 0)
            {
                return;
            }

            if (index > 0)
            {
                values.RemoveAt(index);
                values.Insert(0, target);
                return;
            }

            values.Insert(0, target);
        }

        private static int ClampInt(int value, int min, int max, int fallback)
        {
            if (value == 0)
            {
                value = fallback;
            }
            if (value < min)
            {
                return min;
            }
            if (value > max)
            {
                return max;
            }
            return value;
        }
    }
}
