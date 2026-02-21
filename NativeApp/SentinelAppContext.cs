using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Net.NetworkInformation;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace NetQualitySentinel
{
    internal sealed class SentinelAppContext : ApplicationContext
    {
        private readonly string _appName;
        private readonly string _appDir;
        private readonly string _settingsPath;
        private readonly string _iconPath;
        private readonly object _sync = new object();
        private readonly HttpClient _httpClient;
        private readonly NotifyIcon _notifyIcon;
        private readonly ContextMenuStrip _trayMenu;
        private readonly Queue<double> _downHistory;
        private readonly Queue<double> _upHistory;
        private readonly Queue<double> _successHistory;
        private readonly Queue<QualityHistoryPoint> _qualityHistory;
        private readonly CancellationTokenSource _cts;
        private readonly Dictionary<string, EndpointBackoffState> _downloadEndpointBackoff;
        private readonly Dictionary<string, EndpointBackoffState> _uploadEndpointBackoff;

        private AppConfig _config;
        private Icon _currentIcon;
        private Snapshot _lastSnapshot;
        private LatencyResult _lastLatency;
        private ProbeResult _lastDownProbe;
        private ProbeResult _lastUpProbe;
        private byte[] _uploadBuffer;
        private bool _paused;
        private bool _settingsOpen;
        private bool _forceProbe;
        private DateTime _nextLatencyAtUtc;
        private DateTime _nextSpeedAtUtc;
        private DateTime _nextFullProbeAtUtc;
        private double _lastGoodDownMbps;
        private double _lastGoodUpMbps;
        private DateTime _lastGoodDownAtUtc;
        private DateTime _lastGoodUpAtUtc;
        private string _preferredHost;
        private string _lastIconKey;
        private string _lastTooltip;
        private bool _disposed;
        private StatusForm _statusForm;
        private ToolStripMenuItem _pauseMenuItem;
        private readonly bool _openGraphOnStart;
        private readonly bool _openSettingsOnStart;
        private System.Windows.Forms.Timer _autoOpenGraphTimer;
        private System.Windows.Forms.Timer _autoOpenSettingsTimer;
        private int _autoOpenGraphAttempts;
        private DateTime _nextHistorySampleAtUtc;
        private int _speedProbeCount;

        private const int QualityHistoryWindowSeconds = 60;
        private const int QualityHistorySampleIntervalSec = 1;
        private const int DownloadProbeTrialsPerEndpoint = 3;
        private const int UploadProbeTrialsPerEndpoint = 3;
        private const int DownloadParallelStreamsSmall = 2;
        private const int DownloadParallelStreamsFull = 3;
        private const int UploadParallelStreamsSmall = 2;
        private const int UploadParallelStreamsFull = 3;
        private const int MinParallelDownloadLaneBytes = 512000;
        private const int MinParallelUploadLaneBytes = 384000;
        private const int MaxProbeWarmupBytes = 131072;
        private const int MinReliableProbeBytes = 65536;
        private const int MinReliableProbeDurationMs = 280;
        private const int TargetProbeDurationSmallMs = 1200;
        private const int TargetProbeDurationFullMs = 3000;
        private const int MaxAdaptiveDownloadBytesSmall = 16000000;
        private const int MaxAdaptiveDownloadBytesFull = 48000000;
        private const int MaxAdaptiveUploadBytesSmall = 10000000;
        private const int MaxAdaptiveUploadBytesFull = 28000000;
        private const int MaxUploadBufferBytes = 64000000;
        private const string StartupValueName = "NQA";

        public SentinelAppContext(string appName, string appDir, string settingsPath, string iconPath, bool openGraphOnStart, bool openSettingsOnStart)
        {
            _appName = appName;
            _appDir = appDir;
            _settingsPath = settingsPath;
            _iconPath = iconPath;
            _openGraphOnStart = openGraphOnStart;
            _openSettingsOnStart = openSettingsOnStart;

            _config = AppConfig.Load(_settingsPath);
            _config.Normalize();

            ConfigureHttpNetworking();
            _httpClient = new HttpClient();
            _httpClient.Timeout = TimeSpan.FromSeconds(_config.HttpTimeoutSec);
            TrySetUserAgent();

            _downHistory = new Queue<double>();
            _upHistory = new Queue<double>();
            _successHistory = new Queue<double>();
            _qualityHistory = new Queue<QualityHistoryPoint>();
            _downloadEndpointBackoff = new Dictionary<string, EndpointBackoffState>(StringComparer.OrdinalIgnoreCase);
            _uploadEndpointBackoff = new Dictionary<string, EndpointBackoffState>(StringComparer.OrdinalIgnoreCase);

            _lastLatency = LatencyResult.Empty();
            _lastDownProbe = ProbeResult.Fail("No download probe yet.");
            _lastUpProbe = ProbeResult.Fail("No upload probe yet.");

            _lastGoodDownMbps = 0.0;
            _lastGoodUpMbps = 0.0;
            _lastGoodDownAtUtc = DateTime.UtcNow.AddYears(-10);
            _lastGoodUpAtUtc = DateTime.UtcNow.AddYears(-10);

            _paused = false;
            _settingsOpen = false;
            _forceProbe = true;
            _lastIconKey = string.Empty;
            _lastTooltip = string.Empty;
            _uploadBuffer = CreateUploadBuffer(_config);
            _nextHistorySampleAtUtc = AlignToNextSecond(DateTime.UtcNow);

            _notifyIcon = new NotifyIcon();
            _notifyIcon.Visible = true;
            _currentIcon = BuildStateIcon(0, "Offline", _config);
            _notifyIcon.Icon = _currentIcon;
            _notifyIcon.Text = ClipTooltip(_appName + " (left-click for graph)");
            _trayMenu = BuildTrayMenu();
            _notifyIcon.ContextMenuStrip = _trayMenu;
            _notifyIcon.MouseUp += NotifyIcon_MouseUp;
            _notifyIcon.DoubleClick += NotifyIcon_DoubleClick;
            UpdatePauseMenuItem();
            EnsureSearchShortcut();

            _nextLatencyAtUtc = DateTime.UtcNow;
            _nextSpeedAtUtc = DateTime.UtcNow;
            _nextFullProbeAtUtc = DateTime.UtcNow.AddSeconds(_config.FullProbeIntervalSec);

            if (_openGraphOnStart)
            {
                StartAutoOpenGraphTimer();
            }
            if (_openSettingsOnStart)
            {
                StartAutoOpenSettingsTimer();
            }

            _cts = new CancellationTokenSource();
            Task.Run(async () => await MonitorLoopAsync(_cts.Token));
        }

        private void TrySetUserAgent()
        {
            try
            {
                _httpClient.DefaultRequestHeaders.UserAgent.ParseAdd("NetQualitySentinel/2.0");
            }
            catch
            {
            }
        }

        private static void ConfigureHttpNetworking()
        {
            try
            {
                ServicePointManager.DefaultConnectionLimit = Math.Max(32, ServicePointManager.DefaultConnectionLimit);
            }
            catch
            {
            }

            try
            {
                ServicePointManager.Expect100Continue = false;
            }
            catch
            {
            }

            try
            {
                ServicePointManager.UseNagleAlgorithm = false;
            }
            catch
            {
            }

            try
            {
                ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12;
            }
            catch
            {
            }
        }

        private void EnsureSearchShortcut()
        {
            try
            {
                string exePath = Process.GetCurrentProcess().MainModule.FileName;
                StartMenuRegistration.EnsureShortcut(_appName, exePath, _iconPath);
            }
            catch
            {
            }
        }

        private ContextMenuStrip BuildTrayMenu()
        {
            ContextMenuStrip menu = new ContextMenuStrip();
            menu.ShowImageMargin = false;
            menu.ShowCheckMargin = true;
            menu.Padding = new Padding(2);
            menu.Font = new Font("Segoe UI", 9.5f, FontStyle.Regular, GraphicsUnit.Point);

            ToolStripMenuItem openGraphItem = new ToolStripMenuItem("Open Quality Graph");
            openGraphItem.Click += delegate
            {
                OpenStatusWindow();
            };

            ToolStripMenuItem openSettingsItem = new ToolStripMenuItem("Open Settings");
            openSettingsItem.Click += delegate
            {
                OpenSettingsDialog();
            };

            ToolStripMenuItem probeNowItem = new ToolStripMenuItem("Probe Now");
            probeNowItem.Click += delegate
            {
                ForceProbeNow();
            };

            _pauseMenuItem = new ToolStripMenuItem("Pause Monitoring");
            _pauseMenuItem.Click += delegate
            {
                TogglePauseState();
            };

            ToolStripMenuItem exitItem = new ToolStripMenuItem("Exit");
            exitItem.Click += delegate
            {
                ExitThread();
            };

            menu.Items.Add(openGraphItem);
            menu.Items.Add(openSettingsItem);
            menu.Items.Add(new ToolStripSeparator());
            menu.Items.Add(probeNowItem);
            menu.Items.Add(_pauseMenuItem);
            menu.Items.Add(new ToolStripSeparator());
            menu.Items.Add(exitItem);
            ApplyTrayMenuTheme(menu, UiTheme.GetCurrent());
            menu.Opening += delegate
            {
                UpdatePauseMenuItem();
                ApplyTrayMenuTheme(menu, UiTheme.GetCurrent());
            };

            return menu;
        }

        private static void ApplyTrayMenuTheme(ContextMenuStrip menu, UiTheme theme)
        {
            if (menu == null || theme == null)
            {
                return;
            }

            menu.BackColor = theme.CardBackground;
            menu.ForeColor = theme.TextPrimary;
            menu.RenderMode = ToolStripRenderMode.Professional;
            menu.Renderer = new ToolStripProfessionalRenderer(new TrayMenuColorTable(theme));

            foreach (ToolStripItem item in menu.Items)
            {
                item.BackColor = theme.CardBackground;
                item.ForeColor = theme.TextPrimary;
                item.Font = menu.Font;
                ToolStripSeparator separator = item as ToolStripSeparator;
                if (separator != null)
                {
                    separator.Margin = new Padding(6, 4, 6, 4);
                }
            }
        }

        private sealed class TrayMenuColorTable : ProfessionalColorTable
        {
            private readonly UiTheme _theme;

            public TrayMenuColorTable(UiTheme theme)
            {
                _theme = theme ?? UiTheme.GetCurrent();
            }

            public override Color ToolStripDropDownBackground
            {
                get { return _theme.CardBackground; }
            }

            public override Color MenuBorder
            {
                get { return _theme.CardBorder; }
            }

            public override Color MenuItemBorder
            {
                get { return Blend(_theme.Accent, _theme.CardBorder, 0.34f); }
            }

            public override Color MenuItemSelected
            {
                get { return Blend(_theme.Accent, _theme.CardBackground, _theme.IsDark ? 0.30f : 0.18f); }
            }

            public override Color MenuItemSelectedGradientBegin
            {
                get { return MenuItemSelected; }
            }

            public override Color MenuItemSelectedGradientEnd
            {
                get { return MenuItemSelected; }
            }

            public override Color MenuItemPressedGradientBegin
            {
                get { return Blend(_theme.Accent, _theme.CardBackground, _theme.IsDark ? 0.40f : 0.24f); }
            }

            public override Color MenuItemPressedGradientMiddle
            {
                get { return MenuItemPressedGradientBegin; }
            }

            public override Color MenuItemPressedGradientEnd
            {
                get { return MenuItemPressedGradientBegin; }
            }

            public override Color ImageMarginGradientBegin
            {
                get { return _theme.CardBackground; }
            }

            public override Color ImageMarginGradientMiddle
            {
                get { return _theme.CardBackground; }
            }

            public override Color ImageMarginGradientEnd
            {
                get { return _theme.CardBackground; }
            }

            public override Color CheckBackground
            {
                get { return Blend(_theme.Accent, _theme.CardBackground, _theme.IsDark ? 0.48f : 0.28f); }
            }

            public override Color CheckSelectedBackground
            {
                get { return Blend(_theme.Accent, _theme.CardBackground, _theme.IsDark ? 0.54f : 0.34f); }
            }

            public override Color CheckPressedBackground
            {
                get { return Blend(_theme.Accent, _theme.CardBackground, _theme.IsDark ? 0.58f : 0.38f); }
            }

            public override Color SeparatorDark
            {
                get { return _theme.CardBorder; }
            }

            public override Color SeparatorLight
            {
                get { return _theme.CardBackground; }
            }
        }

        private static Color Blend(Color top, Color bottom, float amount)
        {
            float clamped = Math.Max(0f, Math.Min(1f, amount));
            int r = (int)Math.Round((top.R * clamped) + (bottom.R * (1f - clamped)));
            int g = (int)Math.Round((top.G * clamped) + (bottom.G * (1f - clamped)));
            int b = (int)Math.Round((top.B * clamped) + (bottom.B * (1f - clamped)));
            return Color.FromArgb(255, r, g, b);
        }

        private void ForceProbeNow()
        {
            _forceProbe = true;
            _nextLatencyAtUtc = DateTime.UtcNow;
            _nextSpeedAtUtc = DateTime.UtcNow;
            _nextFullProbeAtUtc = DateTime.UtcNow;
        }

        private void TogglePauseState()
        {
            _paused = !_paused;
            if (!_paused)
            {
                ForceProbeNow();
            }
            UpdatePauseMenuItem();
        }

        private void UpdatePauseMenuItem()
        {
            if (_pauseMenuItem == null)
            {
                return;
            }

            _pauseMenuItem.Checked = _paused;
            _pauseMenuItem.Text = _paused ? "Resume Monitoring" : "Pause Monitoring";
        }

        private void StartAutoOpenGraphTimer()
        {
            if (_autoOpenGraphTimer != null)
            {
                return;
            }

            _autoOpenGraphAttempts = 0;
            _autoOpenGraphTimer = new System.Windows.Forms.Timer();
            _autoOpenGraphTimer.Interval = 750;
            _autoOpenGraphTimer.Tick += AutoOpenGraphTimer_Tick;
            _autoOpenGraphTimer.Start();
        }

        private void AutoOpenGraphTimer_Tick(object sender, EventArgs e)
        {
            _autoOpenGraphAttempts++;

            bool hasSnapshot;
            lock (_sync)
            {
                hasSnapshot = _lastSnapshot != null;
            }

            if (!hasSnapshot && _autoOpenGraphAttempts < 24)
            {
                return;
            }

            StopAutoOpenGraphTimer();
            if (hasSnapshot)
            {
                OpenStatusWindow();
            }
        }

        private void StopAutoOpenGraphTimer()
        {
            if (_autoOpenGraphTimer == null)
            {
                return;
            }

            try { _autoOpenGraphTimer.Stop(); } catch { }
            try { _autoOpenGraphTimer.Tick -= AutoOpenGraphTimer_Tick; } catch { }
            try { _autoOpenGraphTimer.Dispose(); } catch { }
            _autoOpenGraphTimer = null;
        }

        private void StartAutoOpenSettingsTimer()
        {
            if (_autoOpenSettingsTimer != null)
            {
                return;
            }

            _autoOpenSettingsTimer = new System.Windows.Forms.Timer();
            _autoOpenSettingsTimer.Interval = 450;
            _autoOpenSettingsTimer.Tick += AutoOpenSettingsTimer_Tick;
            _autoOpenSettingsTimer.Start();
        }

        private void AutoOpenSettingsTimer_Tick(object sender, EventArgs e)
        {
            StopAutoOpenSettingsTimer();
            OpenSettingsDialog();
        }

        private void StopAutoOpenSettingsTimer()
        {
            if (_autoOpenSettingsTimer == null)
            {
                return;
            }

            try { _autoOpenSettingsTimer.Stop(); } catch { }
            try { _autoOpenSettingsTimer.Tick -= AutoOpenSettingsTimer_Tick; } catch { }
            try { _autoOpenSettingsTimer.Dispose(); } catch { }
            _autoOpenSettingsTimer = null;
        }

        private void NotifyIcon_DoubleClick(object sender, EventArgs e)
        {
            OpenStatusWindow();
        }

        private void NotifyIcon_MouseUp(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                OpenStatusWindow();
            }
        }

        private void OpenSettingsDialog()
        {
            if (_settingsOpen)
            {
                return;
            }

            _settingsOpen = true;
            try
            {
                Snapshot snapshot;
                lock (_sync)
                {
                    snapshot = _lastSnapshot != null ? _lastSnapshot.Clone() : null;
                }

                AppConfig draft = _config.Clone();
                bool startupEnabled = StartupRegistration.IsEnabled(StartupValueName) || StartupRegistration.IsEnabled("NetQualityTray");

                using (SettingsForm form = new SettingsForm(
                    _appName,
                    draft,
                    snapshot,
                    _paused,
                    startupEnabled,
                    _iconPath))
                {
                    form.ShowDialog();

                    if (form.ExitRequested)
                    {
                        ExitThread();
                        return;
                    }

                    if (form.Saved && form.UpdatedConfig != null)
                    {
                        _config = form.UpdatedConfig;
                        _config.Normalize();
                        _paused = form.Paused;
                        ForceProbeNow();
                        _speedProbeCount = 0;
                        _nextFullProbeAtUtc = DateTime.UtcNow.AddSeconds(_config.FullProbeIntervalSec);
                        _httpClient.Timeout = TimeSpan.FromSeconds(_config.HttpTimeoutSec);
                        _uploadBuffer = CreateUploadBuffer(_config);

                        AppConfig.Save(_settingsPath, _config);

                        string exePath = Process.GetCurrentProcess().MainModule.FileName;
                        StartupRegistration.SetEnabled(StartupValueName, exePath, form.StartupEnabled);
                        StartupRegistration.SetEnabled("NetQualityTray", exePath, false);
                        UpdatePauseMenuItem();
                    }

                    if (form.ForceProbe)
                    {
                        ForceProbeNow();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Unable to open settings: " + ex.Message,
                    _appName,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
            }
            finally
            {
                _settingsOpen = false;
            }
        }

        private void OpenStatusWindow()
        {
            Snapshot snapshot;
            lock (_sync)
            {
                snapshot = _lastSnapshot != null ? _lastSnapshot.Clone() : null;
            }

            if (snapshot == null)
            {
                MessageBox.Show(
                    "No samples collected yet.",
                    _appName,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information
                );
                return;
            }

            if (_statusForm == null || _statusForm.IsDisposed)
            {
                _statusForm = new StatusForm(_appName, snapshot, _iconPath);
                _statusForm.FormClosed += delegate
                {
                    _statusForm = null;
                };
                _statusForm.Show();
                return;
            }

            _statusForm.UpdateSnapshot(snapshot);
            if (!_statusForm.Visible)
            {
                _statusForm.Show();
            }
            _statusForm.BringToFront();
            _statusForm.Activate();
        }

        private async Task MonitorLoopAsync(CancellationToken token)
        {
            while (!token.IsCancellationRequested)
            {
                try
                {
                    DateTime now = DateTime.UtcNow;

                    if (_forceProbe)
                    {
                        _nextLatencyAtUtc = now;
                        _nextSpeedAtUtc = now;
                        _forceProbe = false;
                    }

                    InterfaceInfo nic = GetPrimaryInterfaceInfo();

                    if (!_paused)
                    {
                        if (now >= _nextLatencyAtUtc)
                        {
                            _lastLatency = await MeasureLatencyAsync(_config, _preferredHost, token);
                            if (_lastLatency.Success && !string.IsNullOrWhiteSpace(_lastLatency.Host))
                            {
                                _preferredHost = _lastLatency.Host;
                            }
                            _nextLatencyAtUtc = now.AddSeconds(_config.LatencyIntervalSec);
                        }

                        if (now >= _nextSpeedAtUtc)
                        {
                            bool runFull = now >= _nextFullProbeAtUtc || _speedProbeCount < 2;
                            int configuredDownBytes = runFull ? _config.DownloadFullBytes : _config.DownloadSmallBytes;
                            int configuredUpBytes = runFull ? _config.UploadFullBytes : _config.UploadSmallBytes;
                            int downBytes = ResolveDownloadProbeBytes(configuredDownBytes, runFull);
                            int upBytes = ResolveUploadProbeBytes(configuredUpBytes, runFull);

                            ProbeResult downResult = await MeasureDownloadProbeAsync(_config.DownloadEndpoints, downBytes, _config.DownloadProbeMaxMs, runFull, token);
                            ProbeResult upResult = await MeasureUploadProbeAsync(_config.UploadEndpoints, upBytes, _config.UploadProbeMaxMs, runFull, token);
                            _speedProbeCount++;

                            _lastDownProbe = downResult;
                            _lastUpProbe = upResult;

                            if (downResult.Success)
                            {
                                _lastGoodDownMbps = downResult.Mbps;
                                _lastGoodDownAtUtc = now;
                                AddWindowSample(_downHistory, _lastGoodDownMbps, _config.HistorySize);
                            }
                            if (upResult.Success)
                            {
                                _lastGoodUpMbps = upResult.Mbps;
                                _lastGoodUpAtUtc = now;
                                AddWindowSample(_upHistory, _lastGoodUpMbps, _config.HistorySize);
                            }

                            AddWindowSample(_successHistory, (downResult.Success || upResult.Success) ? 1.0 : 0.0, _config.HistorySize);

                            if (runFull)
                            {
                                _nextFullProbeAtUtc = now.AddSeconds(_config.FullProbeIntervalSec);
                            }
                        }
                    }

                    Snapshot snapshot = ComputeSnapshot(nic, now);
                    AddFixedIntervalHistorySamples(snapshot, now);
                    snapshot.QualityHistory = CloneQualityHistory(now);
                    lock (_sync)
                    {
                        _lastSnapshot = snapshot;
                    }
                    UpdateTray(snapshot);
                    UpdateOpenStatusWindow(snapshot);

                    if (_paused)
                    {
                        _nextSpeedAtUtc = now.AddSeconds(5);
                    }
                    else if (now >= _nextSpeedAtUtc)
                    {
                        int intervalSec;
                        if (snapshot.Offline || snapshot.QualityScore < 40.0)
                        {
                            intervalSec = _config.SpeedIntervalPoorSec;
                        }
                        else if (snapshot.QualityScore < 70.0)
                        {
                            intervalSec = _config.SpeedIntervalNormalSec;
                        }
                        else
                        {
                            intervalSec = _config.SpeedIntervalGoodSec;
                        }
                        _nextSpeedAtUtc = now.AddSeconds(intervalSec);
                    }
                }
                catch (TaskCanceledException)
                {
                    break;
                }
                catch
                {
                }

                try
                {
                    await Task.Delay(500, token);
                }
                catch (TaskCanceledException)
                {
                    break;
                }
            }
        }

        private void AddFixedIntervalHistorySamples(Snapshot snapshot, DateTime nowUtc)
        {
            if (snapshot == null)
            {
                return;
            }

            while (_nextHistorySampleAtUtc <= nowUtc)
            {
                _qualityHistory.Enqueue(CreateHistoryPoint(snapshot, _nextHistorySampleAtUtc));
                _nextHistorySampleAtUtc = _nextHistorySampleAtUtc.AddSeconds(QualityHistorySampleIntervalSec);
            }

            PruneHistory(nowUtc);
        }

        private static DateTime AlignToNextSecond(DateTime utc)
        {
            DateTime whole = new DateTime(utc.Year, utc.Month, utc.Day, utc.Hour, utc.Minute, utc.Second, DateTimeKind.Utc);
            if (whole < utc)
            {
                whole = whole.AddSeconds(1);
            }
            return whole;
        }

        private static QualityHistoryPoint CreateHistoryPoint(Snapshot snapshot, DateTime timestampUtc)
        {
            return new QualityHistoryPoint
            {
                TimestampUtc = timestampUtc,
                QualityScore = snapshot.QualityScore,
                Tier = snapshot.Tier,
                DownloadMbps = snapshot.DownloadMbps,
                UploadMbps = snapshot.UploadMbps,
                LatencyMs = snapshot.LatencyMs,
                JitterMs = snapshot.JitterMs,
                LossPct = snapshot.LossPct
            };
        }

        private void PruneHistory(DateTime nowUtc)
        {
            DateTime cutoff = nowUtc.AddSeconds(-QualityHistoryWindowSeconds);
            while (_qualityHistory.Count > 0 && _qualityHistory.Peek().TimestampUtc < cutoff)
            {
                _qualityHistory.Dequeue();
            }
        }

        private List<QualityHistoryPoint> CloneQualityHistory(DateTime nowUtc)
        {
            PruneHistory(nowUtc);

            List<QualityHistoryPoint> result = new List<QualityHistoryPoint>(_qualityHistory.Count);
            foreach (QualityHistoryPoint point in _qualityHistory)
            {
                if (point != null)
                {
                    result.Add(point.Clone());
                }
            }
            return result;
        }

        private void UpdateOpenStatusWindow(Snapshot snapshot)
        {
            if (_statusForm == null || _statusForm.IsDisposed || snapshot == null)
            {
                return;
            }

            try
            {
                Snapshot copy = snapshot.Clone();
                _statusForm.BeginInvoke((Action)(() =>
                {
                    if (_statusForm != null && !_statusForm.IsDisposed)
                    {
                        _statusForm.UpdateSnapshot(copy);
                    }
                }));
            }
            catch
            {
            }
        }

        private Snapshot ComputeSnapshot(InterfaceInfo nic, DateTime nowUtc)
        {
            double downMbps = _lastDownProbe.Success
                ? _lastDownProbe.Mbps
                : DecayRecentValue(_lastGoodDownMbps, (nowUtc - _lastGoodDownAtUtc).TotalSeconds, 45.0);
            double upMbps = _lastUpProbe.Success
                ? _lastUpProbe.Mbps
                : DecayRecentValue(_lastGoodUpMbps, (nowUtc - _lastGoodUpAtUtc).TotalSeconds, 45.0);

            double latencyMs = _lastLatency.Success ? _lastLatency.AvgMs : double.NaN;
            double jitterMs = _lastLatency.Success ? _lastLatency.JitterMs : double.NaN;
            double lossPct = _lastLatency.Success ? _lastLatency.LossPct : 100.0;
            double consistency = GetConsistencyScore(_downHistory, _upHistory, _successHistory);
            bool offline = !nic.IsConnected && !_lastDownProbe.Success && !_lastUpProbe.Success;

            double score;
            string tier;

            if (_paused)
            {
                score = _lastSnapshot != null ? _lastSnapshot.QualityScore : 0.0;
                tier = "Paused";
            }
            else if (offline)
            {
                score = 0.0;
                tier = "Offline";
            }
            else
            {
                double downScore = NormalizeLogScale(downMbps, 200.0);
                double upScore = NormalizeLogScale(upMbps, 80.0);
                double latencyScore = double.IsNaN(latencyMs) ? 0.0 : Clamp(1.0 - (latencyMs / 180.0), 0.0, 1.0);
                double jitterScore = double.IsNaN(jitterMs) ? 0.0 : Clamp(1.0 - (jitterMs / 80.0), 0.0, 1.0);
                double lossScore = Clamp(1.0 - (lossPct / 20.0), 0.0, 1.0);

                score = 100.0 * (
                    0.34 * downScore +
                    0.18 * upScore +
                    0.18 * latencyScore +
                    0.10 * jitterScore +
                    0.10 * lossScore +
                    0.10 * consistency
                );
                score = Clamp(score, 0.0, 100.0);
                tier = ResolveTier(score, offline, _config);
            }

            Snapshot snapshot = new Snapshot();
            snapshot.TimestampUtc = nowUtc;
            snapshot.InterfaceName = nic.Name;
            snapshot.InterfaceType = nic.InterfaceType;
            snapshot.LinkMbps = nic.LinkMbps;
            snapshot.DownloadMbps = Math.Round(downMbps, 2);
            snapshot.UploadMbps = Math.Round(upMbps, 2);
            snapshot.LatencyMs = double.IsNaN(latencyMs) ? double.NaN : Math.Round(latencyMs, 1);
            snapshot.JitterMs = double.IsNaN(jitterMs) ? double.NaN : Math.Round(jitterMs, 1);
            snapshot.LossPct = Math.Round(lossPct, 1);
            snapshot.ConsistencyScore = Clamp(consistency, 0.0, 1.0);
            snapshot.QualityScore = Math.Round(score, 1);
            snapshot.Offline = offline;
            snapshot.Tier = tier;
            snapshot.LastDownloadError = _lastDownProbe.Error ?? string.Empty;
            snapshot.LastUploadError = _lastUpProbe.Error ?? string.Empty;
            snapshot.LatencyHost = _lastLatency.Host ?? "n/a";
            return snapshot;
        }

        private int ResolveDownloadProbeBytes(int configuredBytes, bool fullProbe)
        {
            int requested = Math.Max(32768, configuredBytes);
            double fallbackMbps = fullProbe ? 120.0 : 75.0;
            double estimatedMbps = _lastGoodDownMbps > 0.1 ? _lastGoodDownMbps : fallbackMbps;
            double targetSeconds = (fullProbe ? TargetProbeDurationFullMs : TargetProbeDurationSmallMs) / 1000.0;
            double estimatedBytes = (estimatedMbps * 1000000.0 / 8.0) * targetSeconds;
            int floor = fullProbe ? 512000 : 256000;
            int cap = fullProbe ? MaxAdaptiveDownloadBytesFull : MaxAdaptiveDownloadBytesSmall;
            int adaptive = (int)Math.Round(Clamp(estimatedBytes, floor, cap));
            return Math.Max(requested, adaptive);
        }

        private int ResolveUploadProbeBytes(int configuredBytes, bool fullProbe)
        {
            int requested = Math.Max(4096, configuredBytes);
            double fallbackMbps = fullProbe ? 40.0 : 25.0;
            double estimatedMbps = _lastGoodUpMbps > 0.1 ? _lastGoodUpMbps : fallbackMbps;
            double targetSeconds = (fullProbe ? (TargetProbeDurationFullMs - 300) : (TargetProbeDurationSmallMs - 200)) / 1000.0;
            double estimatedBytes = (estimatedMbps * 1000000.0 / 8.0) * targetSeconds;
            int floor = fullProbe ? 512000 : 256000;
            int cap = fullProbe ? MaxAdaptiveUploadBytesFull : MaxAdaptiveUploadBytesSmall;
            int adaptive = (int)Math.Round(Clamp(estimatedBytes, floor, cap));
            return Math.Max(requested, adaptive);
        }

        private void EnsureUploadBufferCapacity(int requiredBytes)
        {
            int required = Math.Max(8192, requiredBytes);
            if (_uploadBuffer != null && _uploadBuffer.Length >= required)
            {
                return;
            }

            int current = _uploadBuffer != null ? _uploadBuffer.Length : 0;
            int grown = current <= 0 ? required : Math.Max(required, current * 2);
            int newSize = Math.Min(Math.Max(required, grown), Math.Max(required, MaxUploadBufferBytes));
            byte[] buffer = new byte[newSize];
            Random random = new Random();
            random.NextBytes(buffer);
            _uploadBuffer = buffer;
        }

        private static string ResolveTier(double score, bool offline, AppConfig config)
        {
            if (offline)
            {
                return "Offline";
            }
            if (score >= config.QualityHighMinScore)
            {
                return "High";
            }
            if (score >= config.QualityPoorMinScore)
            {
                return "Poor";
            }
            if (score >= config.QualityVeryPoorMinScore)
            {
                return "VeryPoor";
            }
            return "Bad";
        }

        private void UpdateTray(Snapshot snapshot)
        {
            string iconKey = snapshot.Tier + "|" + ((int)Math.Round(snapshot.QualityScore / 5.0, MidpointRounding.AwayFromZero)).ToString();
            string latencyPart = double.IsNaN(snapshot.LatencyMs) ? "n/a" : snapshot.LatencyMs.ToString("0") + " ms";
            string tooltip = string.Format(
                "{0} {1} Q{2:0}% D{3:0.0}Mbps U{4:0.0}Mbps L{5}",
                _appName,
                snapshot.Tier,
                snapshot.QualityScore,
                snapshot.DownloadMbps,
                snapshot.UploadMbps,
                latencyPart
            );

            if (iconKey != _lastIconKey)
            {
                Icon newIcon = BuildStateIcon(snapshot.QualityScore, snapshot.Tier, _config);
                Icon old = _currentIcon;
                _notifyIcon.Icon = newIcon;
                _currentIcon = newIcon;
                _lastIconKey = iconKey;
                if (old != null)
                {
                    old.Dispose();
                }
            }

            if (!string.Equals(_lastTooltip, tooltip, StringComparison.Ordinal))
            {
                _notifyIcon.Text = ClipTooltip(tooltip);
                _lastTooltip = tooltip;
            }
        }

        private static string ClipTooltip(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return "NQA";
            }

            string trimmed = text.Trim();
            if (trimmed.Length <= 63)
            {
                return trimmed;
            }
            return trimmed.Substring(0, 60) + "...";
        }

        private static double NormalizeLogScale(double value, double maxValue)
        {
            if (value <= 0.0 || maxValue <= 0.0)
            {
                return 0.0;
            }
            double numerator = Math.Log10(value + 1.0);
            double denominator = Math.Log10(maxValue + 1.0);
            if (denominator <= 0.0)
            {
                return 0.0;
            }
            return Clamp(numerator / denominator, 0.0, 1.0);
        }

        private static double DecayRecentValue(double value, double ageSeconds, double halfLifeSeconds)
        {
            if (value <= 0.0)
            {
                return 0.0;
            }
            if (ageSeconds <= 0.0)
            {
                return value;
            }

            double decay = Math.Pow(0.5, ageSeconds / Math.Max(halfLifeSeconds, 1.0));
            return value * decay;
        }

        private static double Clamp(double value, double min, double max)
        {
            if (value < min) return min;
            if (value > max) return max;
            return value;
        }

        private static void AddWindowSample(Queue<double> queue, double value, int maxCount)
        {
            if (double.IsNaN(value) || double.IsInfinity(value))
            {
                return;
            }
            queue.Enqueue(value);
            while (queue.Count > maxCount)
            {
                queue.Dequeue();
            }
        }

        private static double Mean(IEnumerable<double> values)
        {
            double sum = 0.0;
            int count = 0;
            foreach (double value in values)
            {
                sum += value;
                count++;
            }
            return count == 0 ? 0.0 : sum / count;
        }

        private static double StdDev(IEnumerable<double> values, double mean)
        {
            double sum = 0.0;
            int count = 0;
            foreach (double value in values)
            {
                double delta = value - mean;
                sum += delta * delta;
                count++;
            }
            return count < 2 ? 0.0 : Math.Sqrt(sum / (count - 1));
        }

        private static double GetConsistencyScore(Queue<double> downHistory, Queue<double> upHistory, Queue<double> successHistory)
        {
            double downMean = Mean(downHistory);
            double downStd = StdDev(downHistory, downMean);
            double upMean = Mean(upHistory);
            double upStd = StdDev(upHistory, upMean);
            double successRate = Mean(successHistory);

            double downConsistency = downMean <= 0.0 ? 0.0 : 1.0 - Clamp(downStd / Math.Max(downMean, 0.001), 0.0, 1.0);
            double upConsistency = upMean <= 0.0 ? 0.0 : 1.0 - Clamp(upStd / Math.Max(upMean, 0.001), 0.0, 1.0);
            return Clamp((downConsistency + upConsistency + successRate) / 3.0, 0.0, 1.0);
        }

        private static InterfaceInfo GetPrimaryInterfaceInfo()
        {
            try
            {
                NetworkInterface[] all = NetworkInterface.GetAllNetworkInterfaces();
                List<NetworkInterface> candidates = new List<NetworkInterface>();

                foreach (NetworkInterface nic in all)
                {
                    if (nic.OperationalStatus != OperationalStatus.Up)
                    {
                        continue;
                    }
                    if (nic.NetworkInterfaceType == NetworkInterfaceType.Loopback ||
                        nic.NetworkInterfaceType == NetworkInterfaceType.Tunnel ||
                        nic.NetworkInterfaceType == NetworkInterfaceType.Unknown)
                    {
                        continue;
                    }
                    candidates.Add(nic);
                }

                if (candidates.Count == 0)
                {
                    return InterfaceInfo.Offline();
                }

                NetworkInterface selected = null;
                long maxScore = long.MinValue;
                foreach (NetworkInterface nic in candidates)
                {
                    bool hasGateway = false;
                    try
                    {
                        hasGateway = nic.GetIPProperties().GatewayAddresses.Count > 0;
                    }
                    catch
                    {
                    }

                    long score = nic.Speed + (hasGateway ? 1000000000000L : 0L);
                    if (score > maxScore)
                    {
                        maxScore = score;
                        selected = nic;
                    }
                }

                if (selected == null)
                {
                    return InterfaceInfo.Offline();
                }

                return new InterfaceInfo
                {
                    IsConnected = true,
                    Name = selected.Name,
                    InterfaceType = selected.NetworkInterfaceType.ToString(),
                    LinkMbps = Math.Max(0.0, selected.Speed / 1000000.0)
                };
            }
            catch
            {
                return InterfaceInfo.Offline();
            }
        }

        private async Task<LatencyResult> MeasureLatencyAsync(AppConfig config, string preferredHost, CancellationToken token)
        {
            List<string> orderedHosts = new List<string>();
            if (!string.IsNullOrWhiteSpace(preferredHost))
            {
                orderedHosts.Add(preferredHost.Trim());
            }

            foreach (string host in config.LatencyHosts)
            {
                if (string.IsNullOrWhiteSpace(host))
                {
                    continue;
                }
                string trimmed = host.Trim();
                if (!orderedHosts.Contains(trimmed))
                {
                    orderedHosts.Add(trimmed);
                }
            }

            if (orderedHosts.Count == 0)
            {
                return LatencyResult.Fail();
            }

            string selectedHost = null;
            List<double> rtts = new List<double>();
            int failures = 0;
            int totalSamples = Math.Max(1, config.LatencySamples);

            using (Ping ping = new Ping())
            {
                foreach (string host in orderedHosts)
                {
                    token.ThrowIfCancellationRequested();
                    try
                    {
                        PingReply reply = await ping.SendPingAsync(host, config.LatencyTimeoutMs);
                        if (reply.Status == IPStatus.Success)
                        {
                            selectedHost = host;
                            rtts.Add(reply.RoundtripTime);
                            break;
                        }
                    }
                    catch
                    {
                    }
                }

                if (string.IsNullOrWhiteSpace(selectedHost))
                {
                    return LatencyResult.Fail();
                }

                for (int i = 1; i < totalSamples; i++)
                {
                    token.ThrowIfCancellationRequested();
                    try
                    {
                        PingReply reply = await ping.SendPingAsync(selectedHost, config.LatencyTimeoutMs);
                        if (reply.Status == IPStatus.Success)
                        {
                            rtts.Add(reply.RoundtripTime);
                        }
                        else
                        {
                            failures++;
                        }
                    }
                    catch
                    {
                        failures++;
                    }
                }
            }

            if (rtts.Count == 0)
            {
                return LatencyResult.Fail();
            }

            double mean = Mean(rtts);
            double std = StdDev(rtts, mean);
            int sent = rtts.Count + failures;
            double lossPct = sent <= 0 ? 100.0 : (failures / (double)sent) * 100.0;

            return new LatencyResult
            {
                Success = true,
                Host = selectedHost,
                AvgMs = Math.Round(mean, 1),
                JitterMs = Math.Round(std, 1),
                LossPct = Math.Round(lossPct, 1)
            };
        }

        private async Task<ProbeResult> MeasureDownloadProbeAsync(List<string> endpoints, int bytesToRead, int maxDurationMs, bool fullProbe, CancellationToken token)
        {
            if (endpoints == null || endpoints.Count == 0)
            {
                return ProbeResult.Fail("No download endpoints configured.");
            }

            ProbeResult lastFailure = ProbeResult.Fail("All download endpoints failed.");
            ProbeResult preferredBest = null;
            ProbeResult fallbackBest = null;
            List<string> candidates = GetCandidateEndpoints(endpoints, _downloadEndpointBackoff, DateTime.UtcNow);
            for (int endpointIndex = 0; endpointIndex < candidates.Count; endpointIndex++)
            {
                string endpoint = candidates[endpointIndex];
                if (string.IsNullOrWhiteSpace(endpoint))
                {
                    continue;
                }

                string trimmed = endpoint.Trim();
                List<ProbeResult> endpointSuccesses = new List<ProbeResult>();
                for (int trial = 0; trial < DownloadProbeTrialsPerEndpoint; trial++)
                {
                    ProbeResult attempt = await MeasureSingleDownloadEndpointAsync(trimmed, bytesToRead, maxDurationMs, fullProbe, token);
                    if (attempt.Success)
                    {
                        MarkEndpointSuccess(_downloadEndpointBackoff, trimmed);
                        endpointSuccesses.Add(attempt);
                        int finalizeTargetMs = fullProbe ? TargetProbeDurationFullMs : TargetProbeDurationSmallMs;
                        if (ShouldFinalizeProbe(endpointSuccesses, bytesToRead, finalizeTargetMs))
                        {
                            break;
                        }
                    }
                    else
                    {
                        MarkEndpointFailure(_downloadEndpointBackoff, trimmed);
                        lastFailure = attempt;
                    }
                }

                ProbeResult endpointResult = AggregateProbeResults(endpointSuccesses, lastFailure, "download");
                if (endpointResult != null && endpointResult.Success)
                {
                    endpointResult.Endpoint = trimmed;
                    if (IsPreferredSpeedEndpoint(trimmed, false))
                    {
                        if (preferredBest == null || endpointResult.Mbps > preferredBest.Mbps)
                        {
                            preferredBest = endpointResult;
                        }
                        if (ShouldAcceptPreferredProbe(endpointResult, _lastGoodDownMbps, false))
                        {
                            return endpointResult;
                        }
                    }
                    if (fallbackBest == null || endpointResult.Mbps > fallbackBest.Mbps)
                    {
                        fallbackBest = endpointResult;
                    }
                }
            }

            ProbeResult selected = SelectBestEndpointProbe(preferredBest, fallbackBest, false);
            return selected ?? fallbackBest ?? preferredBest ?? lastFailure;
        }

        private async Task<ProbeResult> MeasureSingleDownloadEndpointAsync(string endpoint, int bytesToRead, int maxDurationMs, bool fullProbe, CancellationToken token)
        {
            int streamCount = ResolveParallelStreamCount(false, fullProbe);
            if (streamCount <= 1)
            {
                return await MeasureSingleDownloadStreamAsync(endpoint, bytesToRead, maxDurationMs, token);
            }

            int bytesPerStream = ResolveParallelLaneBytes(bytesToRead, streamCount, MinParallelDownloadLaneBytes);
            List<Task<ProbeResult>> tasks = new List<Task<ProbeResult>>(streamCount);
            Stopwatch wallClock = Stopwatch.StartNew();
            for (int i = 0; i < streamCount; i++)
            {
                tasks.Add(MeasureSingleDownloadStreamAsync(endpoint, bytesPerStream, maxDurationMs, token));
            }

            ProbeResult[] attempts = await Task.WhenAll(tasks);
            wallClock.Stop();
            int wallDurationMs = (int)Math.Max(1, wallClock.ElapsedMilliseconds);
            return CombineParallelProbeResults(attempts, endpoint, "download", wallDurationMs);
        }

        private async Task<ProbeResult> MeasureSingleDownloadStreamAsync(string endpoint, int bytesToRead, int maxDurationMs, CancellationToken token)
        {
            string url = endpoint.IndexOf("{bytes}", StringComparison.OrdinalIgnoreCase) >= 0
                ? endpoint.Replace("{bytes}", bytesToRead.ToString())
                : endpoint;

            Stopwatch wallClock = Stopwatch.StartNew();
            Stopwatch transferClock = new Stopwatch();
            long bytesRead = 0;
            long timedBytes = 0;
            int warmupBytes = Math.Min(Math.Max(32768, bytesToRead / 5), MaxProbeWarmupBytes);
            try
            {
                using (CancellationTokenSource linked = CancellationTokenSource.CreateLinkedTokenSource(token))
                {
                    linked.CancelAfter(Math.Max(200, maxDurationMs));
                    using (HttpResponseMessage response = await _httpClient.GetAsync(url, HttpCompletionOption.ResponseHeadersRead, linked.Token))
                    {
                        if (!response.IsSuccessStatusCode)
                        {
                            return ProbeResult.Fail("HTTP " + ((int)response.StatusCode).ToString() + " on " + endpoint);
                        }

                        using (Stream stream = await response.Content.ReadAsStreamAsync())
                        {
                            byte[] buffer = new byte[32768];
                            while (bytesRead < bytesToRead && wallClock.ElapsedMilliseconds < maxDurationMs)
                            {
                                int read = await stream.ReadAsync(buffer, 0, buffer.Length, linked.Token);
                                if (read <= 0)
                                {
                                    break;
                                }

                                bytesRead += read;
                                if (!transferClock.IsRunning && bytesRead >= warmupBytes)
                                {
                                    transferClock.Start();
                                }

                                if (transferClock.IsRunning)
                                {
                                    timedBytes += read;
                                }
                            }
                        }
                    }
                }

                wallClock.Stop();
                if (transferClock.IsRunning)
                {
                    transferClock.Stop();
                }

                if (bytesRead < 20480)
                {
                    return ProbeResult.Fail("Too little data from " + endpoint + " (" + bytesRead.ToString() + " bytes).");
                }

                long measuredBytes = timedBytes >= 16384 ? timedBytes : bytesRead;
                double measuredSeconds = transferClock.Elapsed.TotalSeconds > 0.02
                    ? transferClock.Elapsed.TotalSeconds
                    : wallClock.Elapsed.TotalSeconds;
                int measuredMs = (int)Math.Round(measuredSeconds * 1000.0);
                if (measuredBytes < MinReliableProbeBytes && measuredMs < MinReliableProbeDurationMs)
                {
                    return ProbeResult.Fail("Probe sample from " + endpoint + " was too short for accurate measurement.");
                }

                double seconds = Math.Max(measuredSeconds, 0.001);
                double mbps = ((measuredBytes * 8.0) / 1000000.0) / seconds;
                int durationMs = measuredMs > 0 ? measuredMs : (int)Math.Max(1, wallClock.ElapsedMilliseconds);
                return ProbeResult.SuccessResult(Math.Round(mbps, 2), measuredBytes, durationMs, endpoint);
            }
            catch (TaskCanceledException)
            {
                return ProbeResult.Fail("Timeout on " + endpoint + ".");
            }
            catch (Exception ex)
            {
                return ProbeResult.Fail("Download failed on " + endpoint + ": " + ex.Message);
            }
        }

        private async Task<ProbeResult> MeasureUploadProbeAsync(List<string> endpoints, int bytesToSend, int maxDurationMs, bool fullProbe, CancellationToken token)
        {
            if (endpoints == null || endpoints.Count == 0)
            {
                return ProbeResult.Fail("No upload endpoints configured.");
            }

            ProbeResult lastFailure = ProbeResult.Fail("All upload endpoints failed.");
            ProbeResult preferredBest = null;
            ProbeResult fallbackBest = null;
            List<string> candidates = GetCandidateEndpoints(endpoints, _uploadEndpointBackoff, DateTime.UtcNow);
            for (int endpointIndex = 0; endpointIndex < candidates.Count; endpointIndex++)
            {
                string endpoint = candidates[endpointIndex];
                if (string.IsNullOrWhiteSpace(endpoint))
                {
                    continue;
                }

                string trimmed = endpoint.Trim();
                List<ProbeResult> endpointSuccesses = new List<ProbeResult>();
                for (int trial = 0; trial < UploadProbeTrialsPerEndpoint; trial++)
                {
                    ProbeResult attempt = await MeasureSingleUploadEndpointAsync(trimmed, bytesToSend, maxDurationMs, fullProbe, token);
                    if (attempt.Success)
                    {
                        MarkEndpointSuccess(_uploadEndpointBackoff, trimmed);
                        endpointSuccesses.Add(attempt);
                        int finalizeTargetMs = fullProbe ? (TargetProbeDurationFullMs - 150) : (TargetProbeDurationSmallMs - 150);
                        if (ShouldFinalizeProbe(endpointSuccesses, bytesToSend, finalizeTargetMs))
                        {
                            break;
                        }
                    }
                    else
                    {
                        MarkEndpointFailure(_uploadEndpointBackoff, trimmed);
                        lastFailure = attempt;
                    }
                }

                ProbeResult endpointResult = AggregateProbeResults(endpointSuccesses, lastFailure, "upload");
                if (endpointResult != null && endpointResult.Success)
                {
                    endpointResult.Endpoint = trimmed;
                    if (IsPreferredSpeedEndpoint(trimmed, true))
                    {
                        if (preferredBest == null || endpointResult.Mbps > preferredBest.Mbps)
                        {
                            preferredBest = endpointResult;
                        }
                        if (ShouldAcceptPreferredProbe(endpointResult, _lastGoodUpMbps, true))
                        {
                            return endpointResult;
                        }
                    }
                    if (fallbackBest == null || endpointResult.Mbps > fallbackBest.Mbps)
                    {
                        fallbackBest = endpointResult;
                    }
                }
            }

            ProbeResult selected = SelectBestEndpointProbe(preferredBest, fallbackBest, true);
            return selected ?? fallbackBest ?? preferredBest ?? lastFailure;
        }

        private async Task<ProbeResult> MeasureSingleUploadEndpointAsync(string endpoint, int bytesToSend, int maxDurationMs, bool fullProbe, CancellationToken token)
        {
            int streamCount = ResolveParallelStreamCount(true, fullProbe);
            if (streamCount <= 1)
            {
                return await MeasureSingleUploadStreamAsync(endpoint, bytesToSend, maxDurationMs, token);
            }

            int bytesPerStream = ResolveParallelLaneBytes(bytesToSend, streamCount, MinParallelUploadLaneBytes);
            List<Task<ProbeResult>> tasks = new List<Task<ProbeResult>>(streamCount);
            Stopwatch wallClock = Stopwatch.StartNew();
            for (int i = 0; i < streamCount; i++)
            {
                tasks.Add(MeasureSingleUploadStreamAsync(endpoint, bytesPerStream, maxDurationMs, token));
            }

            ProbeResult[] attempts = await Task.WhenAll(tasks);
            wallClock.Stop();
            int wallDurationMs = (int)Math.Max(1, wallClock.ElapsedMilliseconds);
            return CombineParallelProbeResults(attempts, endpoint, "upload", wallDurationMs);
        }

        private async Task<ProbeResult> MeasureSingleUploadStreamAsync(string endpoint, int bytesToSend, int maxDurationMs, CancellationToken token)
        {
            int requestedBytes = Math.Max(1024, bytesToSend);
            EnsureUploadBufferCapacity(requestedBytes);
            int sendBytes = Math.Min(requestedBytes, _uploadBuffer.Length);
            Stopwatch wallClock = Stopwatch.StartNew();
            Stopwatch transferClock = new Stopwatch();

            try
            {
                using (CancellationTokenSource linked = CancellationTokenSource.CreateLinkedTokenSource(token))
                {
                    linked.CancelAfter(Math.Max(200, maxDurationMs));
                    using (HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Post, endpoint))
                    using (MeasuredUploadContent content = new MeasuredUploadContent(_uploadBuffer, sendBytes, transferClock))
                    {
                        content.Headers.ContentType = new System.Net.Http.Headers.MediaTypeHeaderValue("application/octet-stream");
                        request.Content = content;
                        using (HttpResponseMessage response = await _httpClient.SendAsync(request, HttpCompletionOption.ResponseHeadersRead, linked.Token))
                        {
                            if (!response.IsSuccessStatusCode)
                            {
                                return ProbeResult.Fail("HTTP " + ((int)response.StatusCode).ToString() + " on " + endpoint);
                            }
                        }
                    }
                }

                wallClock.Stop();
                if (transferClock.IsRunning)
                {
                    transferClock.Stop();
                }

                double measuredSeconds = transferClock.Elapsed.TotalSeconds > 0.02
                    ? transferClock.Elapsed.TotalSeconds
                    : wallClock.Elapsed.TotalSeconds;
                int measuredMs = (int)Math.Round(measuredSeconds * 1000.0);
                if (sendBytes < MinReliableProbeBytes && measuredMs < MinReliableProbeDurationMs)
                {
                    return ProbeResult.Fail("Upload sample to " + endpoint + " was too short for accurate measurement.");
                }

                double seconds = Math.Max(measuredSeconds, 0.001);
                double mbps = ((sendBytes * 8.0) / 1000000.0) / seconds;
                int durationMs = measuredMs > 0 ? measuredMs : (int)Math.Max(1, wallClock.ElapsedMilliseconds);
                return ProbeResult.SuccessResult(Math.Round(mbps, 2), sendBytes, durationMs, endpoint);
            }
            catch (TaskCanceledException)
            {
                return ProbeResult.Fail("Timeout on " + endpoint + ".");
            }
            catch (Exception ex)
            {
                return ProbeResult.Fail("Upload failed on " + endpoint + ": " + ex.Message);
            }
        }

        private static List<string> GetCandidateEndpoints(List<string> endpoints, Dictionary<string, EndpointBackoffState> backoffMap, DateTime nowUtc)
        {
            List<string> all = new List<string>();
            HashSet<string> seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (endpoints != null)
            {
                for (int i = 0; i < endpoints.Count; i++)
                {
                    string endpoint = endpoints[i];
                    if (string.IsNullOrWhiteSpace(endpoint))
                    {
                        continue;
                    }

                    string trimmed = endpoint.Trim();
                    if (trimmed.Length == 0 || seen.Contains(trimmed))
                    {
                        continue;
                    }

                    seen.Add(trimmed);
                    all.Add(trimmed);
                }
            }

            if (all.Count == 0)
            {
                return all;
            }

            List<string> available = new List<string>(all.Count);
            for (int i = 0; i < all.Count; i++)
            {
                if (!IsEndpointBackedOff(backoffMap, all[i], nowUtc))
                {
                    available.Add(all[i]);
                }
            }

            return available.Count > 0 ? available : all;
        }

        private static bool IsEndpointBackedOff(Dictionary<string, EndpointBackoffState> backoffMap, string endpoint, DateTime nowUtc)
        {
            if (backoffMap == null || string.IsNullOrWhiteSpace(endpoint))
            {
                return false;
            }

            EndpointBackoffState state;
            if (!backoffMap.TryGetValue(endpoint.Trim(), out state) || state == null)
            {
                return false;
            }

            return state.BackoffUntilUtc > nowUtc;
        }

        private void MarkEndpointSuccess(Dictionary<string, EndpointBackoffState> backoffMap, string endpoint)
        {
            if (backoffMap == null || string.IsNullOrWhiteSpace(endpoint))
            {
                return;
            }

            string key = endpoint.Trim();
            EndpointBackoffState state;
            if (!backoffMap.TryGetValue(key, out state) || state == null)
            {
                return;
            }

            state.ConsecutiveFailures = 0;
            state.BackoffUntilUtc = DateTime.MinValue;
        }

        private void MarkEndpointFailure(Dictionary<string, EndpointBackoffState> backoffMap, string endpoint)
        {
            if (backoffMap == null || string.IsNullOrWhiteSpace(endpoint))
            {
                return;
            }

            string key = endpoint.Trim();
            EndpointBackoffState state;
            if (!backoffMap.TryGetValue(key, out state) || state == null)
            {
                state = new EndpointBackoffState();
                backoffMap[key] = state;
            }

            state.ConsecutiveFailures = Math.Max(0, Math.Min(24, state.ConsecutiveFailures + 1));
            int maxBackoff = _config != null ? _config.MaxEndpointBackoffSec : 180;
            int backoffSec = ComputeBackoffSeconds(state.ConsecutiveFailures, maxBackoff);
            state.BackoffUntilUtc = DateTime.UtcNow.AddSeconds(backoffSec);
        }

        private static int ComputeBackoffSeconds(int failures, int maxBackoffSec)
        {
            int safeMax = Math.Max(10, maxBackoffSec);
            int shift = Math.Max(0, Math.Min(7, failures - 1));
            int baseSeconds = 5 * (1 << shift);
            if (baseSeconds > safeMax)
            {
                baseSeconds = safeMax;
            }
            return Math.Max(5, baseSeconds);
        }

        private static int ResolveParallelLaneBytes(int totalBytes, int streamCount, int minLaneBytes)
        {
            int safeStreams = Math.Max(1, streamCount);
            int safeTotal = Math.Max(minLaneBytes, totalBytes);
            int lane = (int)Math.Ceiling(safeTotal / (double)safeStreams);
            return Math.Max(minLaneBytes, lane);
        }

        private static int ResolveParallelStreamCount(bool upload, bool fullProbe)
        {
            if (upload)
            {
                return fullProbe ? UploadParallelStreamsFull : UploadParallelStreamsSmall;
            }

            return fullProbe ? DownloadParallelStreamsFull : DownloadParallelStreamsSmall;
        }

        private static ProbeResult CombineParallelProbeResults(IList<ProbeResult> attempts, string endpoint, string probeLabel, int wallDurationMs)
        {
            ProbeResult lastFailure = ProbeResult.Fail("All parallel " + (probeLabel ?? "probe") + " streams failed.");
            if (attempts == null || attempts.Count == 0)
            {
                return lastFailure;
            }

            double sumLaneMbps = 0.0;
            long totalBytes = 0;
            int maxDurationMs = 0;
            int successCount = 0;

            for (int i = 0; i < attempts.Count; i++)
            {
                ProbeResult attempt = attempts[i];
                if (attempt == null)
                {
                    continue;
                }

                if (!attempt.Success)
                {
                    lastFailure = attempt;
                    continue;
                }

                if (double.IsNaN(attempt.Mbps) || double.IsInfinity(attempt.Mbps) || attempt.Mbps <= 0.0)
                {
                    continue;
                }

                sumLaneMbps += attempt.Mbps;
                totalBytes += Math.Max(0, attempt.Bytes);
                maxDurationMs = Math.Max(maxDurationMs, Math.Max(1, attempt.DurationMs));
                successCount++;
            }

            if (successCount <= 0)
            {
                return lastFailure;
            }

            if (totalBytes < MinReliableProbeBytes && maxDurationMs < MinReliableProbeDurationMs)
            {
                return ProbeResult.Fail("Parallel " + (probeLabel ?? "probe") + " sample was too short for accurate measurement.");
            }

            int durationMs = Math.Max(1, Math.Max(maxDurationMs, wallDurationMs));
            double mbps = 0.0;
            if (totalBytes > 0 && durationMs > 0)
            {
                double seconds = durationMs / 1000.0;
                mbps = ((totalBytes * 8.0) / 1000000.0) / Math.Max(seconds, 0.001);
            }
            if (mbps <= 0.0 && sumLaneMbps > 0.0)
            {
                mbps = sumLaneMbps;
            }
            string endpointLabel = string.IsNullOrWhiteSpace(endpoint)
                ? ("parallel/" + (probeLabel ?? "probe"))
                : endpoint;
            return ProbeResult.SuccessResult(Math.Round(mbps, 2), totalBytes, durationMs, endpointLabel);
        }

        private static bool ShouldAcceptPreferredProbe(ProbeResult preferred, double lastGoodMbps, bool upload)
        {
            if (preferred == null || !preferred.Success)
            {
                return false;
            }

            double baselineFloor = upload ? 9.0 : 30.0;
            double ratioFloor = lastGoodMbps > 0.1 ? (lastGoodMbps * 0.70) : baselineFloor;
            return preferred.Mbps >= Math.Max(baselineFloor, ratioFloor);
        }

        private static ProbeResult SelectBestEndpointProbe(ProbeResult preferred, ProbeResult bestAny, bool upload)
        {
            bool preferredOk = preferred != null && preferred.Success;
            bool anyOk = bestAny != null && bestAny.Success;
            if (preferredOk && !anyOk)
            {
                return preferred;
            }
            if (!preferredOk && anyOk)
            {
                return bestAny;
            }
            if (!preferredOk && !anyOk)
            {
                return null;
            }

            // Prefer the primary endpoint unless an alternate endpoint is dramatically better.
            double requiredGain = upload ? 1.60 : 1.45;
            if (bestAny.Mbps >= (preferred.Mbps * requiredGain))
            {
                return bestAny;
            }

            return preferred;
        }

        private static bool IsPreferredSpeedEndpoint(string endpoint, bool upload)
        {
            if (string.IsNullOrWhiteSpace(endpoint))
            {
                return false;
            }

            string text = endpoint.Trim().ToLowerInvariant();
            if (upload)
            {
                return text.Contains("speed.cloudflare.com/__up");
            }

            return text.Contains("speed.cloudflare.com/__down");
        }

        private static bool ShouldFinalizeProbe(List<ProbeResult> successes, int requestedBytes, int targetDurationMs)
        {
            if (successes == null || successes.Count == 0)
            {
                return false;
            }

            if (successes.Count >= 3)
            {
                return true;
            }

            if (successes.Count < 2)
            {
                return false;
            }

            int totalMs = 0;
            long totalBytes = 0;
            for (int i = 0; i < successes.Count; i++)
            {
                ProbeResult attempt = successes[i];
                if (attempt == null || !attempt.Success)
                {
                    continue;
                }

                totalMs += Math.Max(0, attempt.DurationMs);
                totalBytes += Math.Max(0, attempt.Bytes);
            }

            if (totalMs >= Math.Max(350, targetDurationMs))
            {
                return true;
            }

            long targetBytes = Math.Max(131072, requestedBytes);
            return totalBytes >= (targetBytes * 2L);
        }

        private static ProbeResult AggregateProbeResults(List<ProbeResult> successes, ProbeResult lastFailure, string probeLabel)
        {
            if (successes == null || successes.Count == 0)
            {
                return lastFailure;
            }

            List<ProbeResult> reliable = new List<ProbeResult>();
            for (int i = 0; i < successes.Count; i++)
            {
                ProbeResult attempt = successes[i];
                if (attempt == null || !attempt.Success)
                {
                    continue;
                }

                bool hasReliableBytes = attempt.Bytes >= MinReliableProbeBytes;
                bool hasReliableDuration = attempt.DurationMs >= MinReliableProbeDurationMs;
                if (hasReliableBytes || hasReliableDuration)
                {
                    reliable.Add(attempt);
                }
            }

            if (reliable.Count == 0)
            {
                reliable = new List<ProbeResult>(successes);
            }

            if (reliable.Count >= 4)
            {
                List<ProbeResult> noOutliers = FilterProbeOutliers(reliable);
                if (noOutliers.Count > 0)
                {
                    reliable = noOutliers;
                }
            }

            ProbeResult representative = SelectRepresentativeProbeResult(reliable);
            if (representative == null || !representative.Success)
            {
                ProbeResult fallback = reliable[0];
                if (fallback != null && fallback.Success)
                {
                    return fallback;
                }
                return lastFailure ?? ProbeResult.Fail("Probe failed.");
            }

            long bytesValue = Math.Max(0, representative.Bytes);
            int durationMs = Math.Max(1, representative.DurationMs);
            string endpointLabel = string.IsNullOrWhiteSpace(representative.Endpoint)
                ? ("aggregate/" + (probeLabel ?? "probe"))
                : representative.Endpoint;
            return ProbeResult.SuccessResult(Math.Round(representative.Mbps, 2), bytesValue, durationMs, endpointLabel);
        }

        private static ProbeResult SelectRepresentativeProbeResult(List<ProbeResult> attempts)
        {
            if (attempts == null || attempts.Count == 0)
            {
                return null;
            }

            List<ProbeResult> valid = new List<ProbeResult>();
            for (int i = 0; i < attempts.Count; i++)
            {
                ProbeResult attempt = attempts[i];
                if (attempt != null && attempt.Success && !double.IsNaN(attempt.Mbps) && !double.IsInfinity(attempt.Mbps) && attempt.Mbps > 0.0)
                {
                    valid.Add(attempt);
                }
            }

            if (valid.Count == 0)
            {
                return null;
            }

            valid.Sort(delegate(ProbeResult left, ProbeResult right)
            {
                return left.Mbps.CompareTo(right.Mbps);
            });

            // Use median throughput to avoid over-reporting on short-lived spikes.
            int index = (int)Math.Round((valid.Count - 1) * 0.50, MidpointRounding.AwayFromZero);
            index = Math.Max(0, Math.Min(valid.Count - 1, index));
            return valid[index];
        }

        private static List<ProbeResult> FilterProbeOutliers(List<ProbeResult> attempts)
        {
            List<ProbeResult> sorted = new List<ProbeResult>();
            if (attempts == null || attempts.Count == 0)
            {
                return sorted;
            }

            for (int i = 0; i < attempts.Count; i++)
            {
                ProbeResult attempt = attempts[i];
                if (attempt != null && attempt.Success && !double.IsNaN(attempt.Mbps) && !double.IsInfinity(attempt.Mbps))
                {
                    sorted.Add(attempt);
                }
            }

            if (sorted.Count < 4)
            {
                return sorted;
            }

            sorted.Sort(delegate(ProbeResult left, ProbeResult right)
            {
                return left.Mbps.CompareTo(right.Mbps);
            });

            List<double> values = new List<double>(sorted.Count);
            for (int i = 0; i < sorted.Count; i++)
            {
                values.Add(sorted[i].Mbps);
            }

            double q1 = Percentile(values, 0.25);
            double q3 = Percentile(values, 0.75);
            double iqr = q3 - q1;
            if (iqr <= 0.0)
            {
                return sorted;
            }

            double min = Math.Max(0.0, q1 - (1.5 * iqr));
            double max = q3 + (1.5 * iqr);
            List<ProbeResult> filtered = new List<ProbeResult>();
            for (int i = 0; i < sorted.Count; i++)
            {
                double value = sorted[i].Mbps;
                if (value >= min && value <= max)
                {
                    filtered.Add(sorted[i]);
                }
            }

            return filtered.Count > 0 ? filtered : sorted;
        }

        private static double Percentile(List<double> sortedValues, double percentile)
        {
            if (sortedValues == null || sortedValues.Count == 0)
            {
                return 0.0;
            }

            if (sortedValues.Count == 1)
            {
                return sortedValues[0];
            }

            double clamped = Clamp(percentile, 0.0, 1.0);
            double position = (sortedValues.Count - 1) * clamped;
            int leftIndex = (int)Math.Floor(position);
            int rightIndex = (int)Math.Ceiling(position);
            if (leftIndex == rightIndex)
            {
                return sortedValues[leftIndex];
            }

            double t = position - leftIndex;
            double left = sortedValues[leftIndex];
            double right = sortedValues[rightIndex];
            return left + ((right - left) * t);
        }

        private sealed class MeasuredUploadContent : HttpContent
        {
            private const int ChunkBytes = 32768;
            private readonly byte[] _buffer;
            private readonly int _count;
            private readonly Stopwatch _stopwatch;

            public MeasuredUploadContent(byte[] buffer, int count, Stopwatch stopwatch)
            {
                _buffer = buffer ?? new byte[0];
                _count = Math.Max(0, Math.Min(count, _buffer.Length));
                _stopwatch = stopwatch ?? new Stopwatch();
            }

            protected override async Task SerializeToStreamAsync(Stream stream, TransportContext context)
            {
                int offset = 0;
                _stopwatch.Reset();
                _stopwatch.Start();
                while (offset < _count)
                {
                    int write = Math.Min(ChunkBytes, _count - offset);
                    await stream.WriteAsync(_buffer, offset, write);
                    offset += write;
                }
                await stream.FlushAsync();
                _stopwatch.Stop();
            }

            protected override bool TryComputeLength(out long length)
            {
                length = _count;
                return true;
            }
        }

        private sealed class EndpointBackoffState
        {
            public int ConsecutiveFailures { get; set; }
            public DateTime BackoffUntilUtc { get; set; }
        }

        private static byte[] CreateUploadBuffer(AppConfig config)
        {
            int size = Math.Max(config.UploadSmallBytes, config.UploadFullBytes);
            size = Math.Max(size, 8192);
            byte[] buffer = new byte[size];
            Random random = new Random();
            random.NextBytes(buffer);
            return buffer;
        }

        private static Icon BuildStateIcon(double qualityScore, string tier, AppConfig config)
        {
            int size = 32;
            Color tierColor = config.GetTierColor(tier);

            int bars;
            switch ((tier ?? string.Empty).Trim())
            {
                case "High":
                    bars = 4;
                    break;
                case "Poor":
                    bars = 3;
                    break;
                case "VeryPoor":
                    bars = 2;
                    break;
                case "Bad":
                    bars = 1;
                    break;
                case "Offline":
                    bars = 0;
                    break;
                case "Paused":
                    bars = 4;
                    break;
                default:
                    bars = Math.Max(1, Math.Min(4, (int)Math.Round(qualityScore / 25.0, MidpointRounding.AwayFromZero)));
                    break;
            }

            using (Bitmap bmp = new Bitmap(size, size))
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.Clear(Color.Transparent);

                Rectangle bgRect = new Rectangle(1, 1, 30, 30);
                using (GraphicsPath bgPath = RoundedRect(bgRect, 8))
                using (SolidBrush bgBrush = new SolidBrush(Color.FromArgb(26, 30, 36)))
                using (Pen borderPen = new Pen(Color.FromArgb(80, 94, 110), 1))
                {
                    g.FillPath(bgBrush, bgPath);
                    g.DrawPath(borderPen, bgPath);
                }

                int[] heights = new int[] { 5, 9, 13, 18 };
                for (int i = 0; i < heights.Length; i++)
                {
                    int x = 7 + (i * 5);
                    int h = heights[i];
                    int y = 24 - h;
                    Rectangle barRect = new Rectangle(x, y, 4, h);
                    Color color = i < bars ? tierColor : Color.FromArgb(64, 72, 82);
                    using (SolidBrush barBrush = new SolidBrush(color))
                    {
                        g.FillRectangle(barBrush, barRect);
                    }
                }

                if (string.Equals(tier, "Paused", StringComparison.OrdinalIgnoreCase))
                {
                    using (SolidBrush pauseBrush = new SolidBrush(Color.FromArgb(230, 230, 230)))
                    {
                        g.FillRectangle(pauseBrush, new Rectangle(20, 8, 3, 8));
                        g.FillRectangle(pauseBrush, new Rectangle(25, 8, 3, 8));
                    }
                }

                IntPtr handle = bmp.GetHicon();
                try
                {
                    Icon icon = Icon.FromHandle(handle);
                    return (Icon)icon.Clone();
                }
                finally
                {
                    NativeMethods.DestroyIcon(handle);
                }
            }
        }

        private static GraphicsPath RoundedRect(Rectangle bounds, int radius)
        {
            int diameter = radius * 2;
            GraphicsPath path = new GraphicsPath();
            path.AddArc(bounds.Left, bounds.Top, diameter, diameter, 180, 90);
            path.AddArc(bounds.Right - diameter, bounds.Top, diameter, diameter, 270, 90);
            path.AddArc(bounds.Right - diameter, bounds.Bottom - diameter, diameter, diameter, 0, 90);
            path.AddArc(bounds.Left, bounds.Bottom - diameter, diameter, diameter, 90, 90);
            path.CloseFigure();
            return path;
        }

        protected override void ExitThreadCore()
        {
            if (_disposed)
            {
                base.ExitThreadCore();
                return;
            }

            _disposed = true;

            try { _cts.Cancel(); } catch { }
            try { StopAutoOpenGraphTimer(); } catch { }
            try { StopAutoOpenSettingsTimer(); } catch { }
            try { _notifyIcon.Visible = false; } catch { }
            try { if (_trayMenu != null) { _trayMenu.Dispose(); } } catch { }
            try { _notifyIcon.Dispose(); } catch { }
            try { if (_statusForm != null && !_statusForm.IsDisposed) { _statusForm.Close(); _statusForm.Dispose(); } } catch { }
            try { if (_currentIcon != null) { _currentIcon.Dispose(); } } catch { }
            try { _httpClient.Dispose(); } catch { }
            try { _cts.Dispose(); } catch { }

            base.ExitThreadCore();
        }
    }
}
