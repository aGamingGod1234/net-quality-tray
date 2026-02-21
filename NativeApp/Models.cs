using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace NetQualitySentinel
{
    internal sealed class QualityHistoryPoint
    {
        public DateTime TimestampUtc { get; set; }
        public double QualityScore { get; set; }
        public string Tier { get; set; }
        public double DownloadMbps { get; set; }
        public double UploadMbps { get; set; }
        public double LatencyMs { get; set; }
        public double JitterMs { get; set; }
        public double LossPct { get; set; }

        public QualityHistoryPoint Clone()
        {
            return (QualityHistoryPoint)MemberwiseClone();
        }
    }

    internal sealed class Snapshot
    {
        public DateTime TimestampUtc { get; set; }
        public string InterfaceName { get; set; }
        public string InterfaceType { get; set; }
        public double LinkMbps { get; set; }
        public double DownloadMbps { get; set; }
        public double UploadMbps { get; set; }
        public double LatencyMs { get; set; }
        public double JitterMs { get; set; }
        public double LossPct { get; set; }
        public double ConsistencyScore { get; set; }
        public double QualityScore { get; set; }
        public bool Offline { get; set; }
        public string Tier { get; set; }
        public string LastDownloadError { get; set; }
        public string LastUploadError { get; set; }
        public string LatencyHost { get; set; }
        public List<QualityHistoryPoint> QualityHistory { get; set; }

        public Snapshot Clone()
        {
            Snapshot clone = (Snapshot)MemberwiseClone();
            clone.QualityHistory = new List<QualityHistoryPoint>();
            if (QualityHistory != null)
            {
                foreach (QualityHistoryPoint point in QualityHistory)
                {
                    if (point != null)
                    {
                        clone.QualityHistory.Add(point.Clone());
                    }
                }
            }
            return clone;
        }
    }

    internal sealed class InterfaceInfo
    {
        public bool IsConnected { get; set; }
        public string Name { get; set; }
        public string InterfaceType { get; set; }
        public double LinkMbps { get; set; }

        public static InterfaceInfo Offline()
        {
            return new InterfaceInfo
            {
                IsConnected = false,
                Name = "Offline",
                InterfaceType = "n/a",
                LinkMbps = 0.0
            };
        }
    }

    internal sealed class LatencyResult
    {
        public bool Success { get; set; }
        public string Host { get; set; }
        public double AvgMs { get; set; }
        public double JitterMs { get; set; }
        public double LossPct { get; set; }

        public static LatencyResult Empty()
        {
            return new LatencyResult
            {
                Success = false,
                Host = null,
                AvgMs = double.NaN,
                JitterMs = double.NaN,
                LossPct = 100.0
            };
        }

        public static LatencyResult Fail()
        {
            return Empty();
        }
    }

    internal sealed class ProbeResult
    {
        public bool Success { get; set; }
        public double Mbps { get; set; }
        public long Bytes { get; set; }
        public int DurationMs { get; set; }
        public string Error { get; set; }
        public string Endpoint { get; set; }

        public static ProbeResult SuccessResult(double mbps, long bytes, int durationMs, string endpoint)
        {
            return new ProbeResult
            {
                Success = true,
                Mbps = mbps,
                Bytes = bytes,
                DurationMs = durationMs,
                Error = string.Empty,
                Endpoint = endpoint
            };
        }

        public static ProbeResult Fail(string error)
        {
            return new ProbeResult
            {
                Success = false,
                Mbps = 0.0,
                Bytes = 0,
                DurationMs = 0,
                Error = error ?? "Probe failed.",
                Endpoint = string.Empty
            };
        }
    }

    internal static class NativeMethods
    {
        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern bool DestroyIcon(IntPtr handle);
    }
}
