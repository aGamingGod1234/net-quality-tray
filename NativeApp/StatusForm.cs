
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Win32;

namespace NetQualitySentinel
{
    internal sealed class StatusForm : Form
    {
        private readonly string _appName;
        private UiTheme _theme;
        private string _currentTier;
        private List<QualityHistoryPoint> _history;
        private List<TimelineSample> _timelineCache;
        private DateTime _timelineCacheBuiltAtUtc;
        private int _timelineCacheHistoryVersion;
        private int _historyVersion;

        private Label _lblTier;
        private Label _lblScore;
        private Label _lblDownload;
        private Label _lblUpload;
        private Label _lblLatency;
        private Label _lblJitter;
        private Label _lblLoss;
        private Label _lblInterface;
        private Label _lblUpdated;
        private Panel _badge;
        private Button _btnClose;

        private GraphPanel _mainGraph;
        private GraphPanel _downloadGraph;
        private GraphPanel _uploadGraph;
        private GraphPanel _latencyGraph;
        private GraphPanel _jitterGraph;
        private GraphPanel _lossGraph;
        private TableLayoutPanel _rootLayout;

        private ToolTip _graphToolTip;
        private Control _activeTooltipControl;
        private string _activeTooltipText;
        private Timer _renderTimer;

        private const int WindowSeconds = 60;
        private const int WsExToolWindow = 0x00000080;
        private const int WsExAppWindow = 0x00040000;
        private const uint SwpNoZOrder = 0x0004;
        private const uint SwpNoActivate = 0x0010;
        private const int HeaderRowHeight = 116;
        private const int FooterRowHeight = 64;
        private const int MinGraphRowHeight = 160;
        private const int MaxGraphRowHeight = 360;
        private const int MinMetricsRowHeight = 160;
        private const double TimelineStepSeconds = 0.2;
        private const int TimelineCacheMaxAgeMs = 140;

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int x, int y, int cx, int cy, uint flags);

        protected override CreateParams CreateParams
        {
            get
            {
                CreateParams cp = base.CreateParams;
                cp.ExStyle &= ~WsExToolWindow;
                cp.ExStyle |= WsExAppWindow;
                return cp;
            }
        }

        public StatusForm(string appName, Snapshot snapshot, string iconPath)
        {
            _appName = appName;
            _theme = UiTheme.GetCurrent();
            _currentTier = "Offline";
            _history = new List<QualityHistoryPoint>();
            _timelineCache = new List<TimelineSample>();
            _timelineCacheBuiltAtUtc = DateTime.MinValue;
            _timelineCacheHistoryVersion = -1;
            _historyVersion = 0;
            _activeTooltipControl = null;
            _activeTooltipText = string.Empty;

            BuildWindow(iconPath);
            BuildLayout();
            InitializeGraphTooltips();
            UpdateSnapshot(snapshot);
            ApplyTheme();

            _renderTimer = new Timer();
            _renderTimer.Interval = 250;
            _renderTimer.Tick += RenderTimer_Tick;
            _renderTimer.Start();

            Shown += StatusForm_Shown;
            Resize += StatusForm_Resize;
            SystemEvents.UserPreferenceChanged += SystemEvents_UserPreferenceChanged;
            FormClosed += StatusForm_FormClosed;
        }

        public void UpdateSnapshot(Snapshot snapshot)
        {
            if (snapshot == null)
            {
                _lblTier.Text = "Waiting for samples";
                _lblScore.Text = "Score --";
                _lblDownload.Text = "-- Mbps";
                _lblUpload.Text = "-- Mbps";
                _lblLatency.Text = "n/a";
                _lblJitter.Text = "n/a";
                _lblLoss.Text = "n/a";
                _lblInterface.Text = "n/a";
                _lblUpdated.Text = "Updated: n/a";
                _badge.BackColor = Color.FromArgb(150, 156, 166);
                _history = new List<QualityHistoryPoint>();
                InvalidateTimelineCache();
                InvalidateGraphs();
                return;
            }

            _currentTier = string.IsNullOrWhiteSpace(snapshot.Tier) ? "Offline" : snapshot.Tier;

            _lblTier.Text = _currentTier;
            _lblScore.Text = "Score " + snapshot.QualityScore.ToString("0") + "%";
            _lblDownload.Text = snapshot.DownloadMbps.ToString("0.00") + " Mbps";
            _lblUpload.Text = snapshot.UploadMbps.ToString("0.00") + " Mbps";
            _lblLatency.Text = double.IsNaN(snapshot.LatencyMs) ? "n/a" : snapshot.LatencyMs.ToString("0.0") + " ms";
            _lblJitter.Text = double.IsNaN(snapshot.JitterMs) ? "n/a" : snapshot.JitterMs.ToString("0.0") + " ms";
            _lblLoss.Text = snapshot.LossPct.ToString("0.0") + "%";
            _lblInterface.Text = snapshot.InterfaceName + " (" + snapshot.InterfaceType + ")";
            _lblUpdated.Text = "Updated: " + snapshot.TimestampUtc.ToLocalTime().ToString("h:mm:ss tt");
            _badge.BackColor = GetTierColor(_currentTier);

            _history = new List<QualityHistoryPoint>();
            if (snapshot.QualityHistory != null)
            {
                for (int i = 0; i < snapshot.QualityHistory.Count; i++)
                {
                    if (snapshot.QualityHistory[i] != null)
                    {
                        _history.Add(snapshot.QualityHistory[i].Clone());
                    }
                }
            }

            if (_history.Count == 0)
            {
                _history.Add(new QualityHistoryPoint
                {
                    TimestampUtc = snapshot.TimestampUtc,
                    QualityScore = snapshot.QualityScore,
                    Tier = snapshot.Tier,
                    DownloadMbps = snapshot.DownloadMbps,
                    UploadMbps = snapshot.UploadMbps,
                    LatencyMs = snapshot.LatencyMs,
                    JitterMs = snapshot.JitterMs,
                    LossPct = snapshot.LossPct
                });
            }

            _historyVersion++;
            InvalidateTimelineCache();
            InvalidateGraphs();
        }

        private void BuildWindow(string iconPath)
        {
            Text = _appName + " Quality Graph";
            StartPosition = FormStartPosition.CenterScreen;
            Rectangle workArea = Screen.PrimaryScreen != null
                ? Screen.PrimaryScreen.WorkingArea
                : new Rectangle(0, 0, 1366, 768);
            int width = Math.Max(820, Math.Min(1000, workArea.Width - 48));
            int height = Math.Max(640, Math.Min(760, workArea.Height - 48));
            int left = workArea.Left + Math.Max(0, (workArea.Width - width) / 2);
            int top = workArea.Top + Math.Max(0, (workArea.Height - height) / 2);
            ApplyWindowBounds(left, top, width, height);
            MinimumSize = new Size(560, 420);
            AutoScaleMode = AutoScaleMode.Dpi;
            FormBorderStyle = FormBorderStyle.Sizable;
            ShowInTaskbar = true;
            MinimizeBox = true;
            MaximizeBox = true;
            ShowIcon = true;
            Font = new Font("Segoe UI", 10f, FontStyle.Regular, GraphicsUnit.Point);
            HandleCreated += delegate { NativeUi.ApplyWindowTheme(this, _theme.IsDark); };

            if (File.Exists(iconPath))
            {
                try { Icon = new Icon(iconPath); } catch { }
            }
        }

        private void StatusForm_Shown(object sender, EventArgs e)
        {
            WindowState = FormWindowState.Normal;
            FitToWorkingArea();
            ApplyResponsiveRowHeights();
        }

        private void StatusForm_Resize(object sender, EventArgs e)
        {
            ApplyResponsiveRowHeights();
        }

        private void FitToWorkingArea()
        {
            Rectangle workArea = Screen.FromControl(this).WorkingArea;
            int maxWidth = Math.Max(640, workArea.Width - 24);
            int maxHeight = Math.Max(520, workArea.Height - 24);
            int targetWidth = Math.Min(maxWidth, Math.Max(820, Math.Min(1000, maxWidth)));
            int targetHeight = Math.Min(maxHeight, Math.Max(640, Math.Min(760, maxHeight)));

            int left = workArea.Left + Math.Max(0, (workArea.Width - targetWidth) / 2);
            int top = workArea.Top + Math.Max(0, (workArea.Height - targetHeight) / 2);

            ApplyWindowBounds(left, top, targetWidth, targetHeight);
        }

        private void ApplyWindowBounds(int left, int top, int width, int height)
        {
            if (IsHandleCreated)
            {
                if (!SetWindowPos(Handle, IntPtr.Zero, left, top, width, height, SwpNoZOrder | SwpNoActivate))
                {
                    SetDesktopBounds(left, top, width, height);
                }
                return;
            }

            SetDesktopBounds(left, top, width, height);
        }

        private void ApplyResponsiveRowHeights()
        {
            if (_rootLayout == null || _rootLayout.RowStyles.Count < 4)
            {
                return;
            }

            int dynamicHeaderHeight = HeaderRowHeight;
            if (_lblTier != null && _lblScore != null)
            {
                dynamicHeaderHeight = Math.Max(HeaderRowHeight, _lblTier.PreferredHeight + _lblScore.PreferredHeight + 38);
            }

            int closeHeight = _btnClose != null ? _btnClose.Height : 30;
            int updatedHeight = _lblUpdated != null ? Math.Max(20, _lblUpdated.Font.Height + 6) : 20;
            int dynamicFooterHeight = Math.Max(FooterRowHeight, Math.Max(closeHeight, updatedHeight) + 20);
            dynamicFooterHeight = Math.Min(88, dynamicFooterHeight);

            int contentHeight = _rootLayout.ClientSize.Height - dynamicHeaderHeight - dynamicFooterHeight;
            if (contentHeight <= 0)
            {
                return;
            }

            int graphTarget = (int)Math.Round(contentHeight * 0.56);
            graphTarget = Math.Max(MinGraphRowHeight, Math.Min(MaxGraphRowHeight, graphTarget));
            int metricsTarget = contentHeight - graphTarget;

            if (metricsTarget < MinMetricsRowHeight)
            {
                metricsTarget = MinMetricsRowHeight;
                graphTarget = contentHeight - metricsTarget;
            }
            if (graphTarget < MinGraphRowHeight)
            {
                graphTarget = MinGraphRowHeight;
                metricsTarget = contentHeight - graphTarget;
            }
            if (metricsTarget < 0)
            {
                metricsTarget = 0;
            }

            _rootLayout.RowStyles[0].SizeType = SizeType.Absolute;
            _rootLayout.RowStyles[0].Height = dynamicHeaderHeight;
            _rootLayout.RowStyles[1].SizeType = SizeType.Absolute;
            _rootLayout.RowStyles[1].Height = graphTarget;
            _rootLayout.RowStyles[2].SizeType = SizeType.Absolute;
            _rootLayout.RowStyles[2].Height = metricsTarget;
            _rootLayout.RowStyles[3].SizeType = SizeType.Absolute;
            _rootLayout.RowStyles[3].Height = dynamicFooterHeight;
        }

        private void BuildLayout()
        {
            _rootLayout = new TableLayoutPanel();
            _rootLayout.Dock = DockStyle.Fill;
            _rootLayout.Padding = new Padding(12);
            _rootLayout.ColumnCount = 1;
            _rootLayout.RowCount = 4;
            _rootLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
            _rootLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, HeaderRowHeight));
            _rootLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 320f));
            _rootLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 260f));
            _rootLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, FooterRowHeight));
            Controls.Add(_rootLayout);

            Panel header = CreateCard();
            header.Dock = DockStyle.Fill;
            _rootLayout.Controls.Add(header, 0, 0);

            TableLayoutPanel headerLayout = new TableLayoutPanel();
            headerLayout.Dock = DockStyle.Fill;
            headerLayout.Padding = new Padding(12, 10, 12, 8);
            headerLayout.ColumnCount = 2;
            headerLayout.RowCount = 1;
            headerLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 20f));
            headerLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
            header.Controls.Add(headerLayout);

            _badge = new Panel();
            _badge.Dock = DockStyle.Fill;
            _badge.Margin = new Padding(0, 6, 8, 6);
            headerLayout.Controls.Add(_badge, 0, 0);

            TableLayoutPanel headerText = new TableLayoutPanel();
            headerText.Dock = DockStyle.Fill;
            headerText.BackColor = Color.Transparent;
            headerText.RowCount = 2;
            headerText.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            headerText.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            headerLayout.Controls.Add(headerText, 1, 0);

            _lblTier = new Label();
            _lblTier.AutoSize = true;
            _lblTier.Font = new Font("Segoe UI Variable Text", 15f, FontStyle.Bold, GraphicsUnit.Point);
            headerText.Controls.Add(_lblTier, 0, 0);

            _lblScore = new Label();
            _lblScore.AutoSize = true;
            _lblScore.Tag = "muted";
            _lblScore.Margin = new Padding(0, 2, 0, 0);
            headerText.Controls.Add(_lblScore, 0, 1);

            Panel graphCard = CreateCard();
            graphCard.Dock = DockStyle.Fill;
            _rootLayout.Controls.Add(graphCard, 0, 1);

            TableLayoutPanel graphLayout = new TableLayoutPanel();
            graphLayout.Dock = DockStyle.Fill;
            graphLayout.Padding = new Padding(12, 8, 12, 8);
            graphLayout.RowCount = 2;
            graphLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            graphLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));
            graphCard.Controls.Add(graphLayout);

            TableLayoutPanel graphHeader = new TableLayoutPanel();
            graphHeader.Dock = DockStyle.Top;
            graphHeader.AutoSize = true;
            graphHeader.ColumnCount = 1;
            graphHeader.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
            graphLayout.Controls.Add(graphHeader, 0, 0);

            Label graphTitle = new Label();
            graphTitle.AutoSize = true;
            graphTitle.Font = new Font("Segoe UI Variable Text", 12f, FontStyle.Bold, GraphicsUnit.Point);
            graphTitle.Text = "Past Minute Quality";
            graphHeader.Controls.Add(graphTitle, 0, 0);

            _mainGraph = new GraphPanel();
            _mainGraph.Dock = DockStyle.Fill;
            _mainGraph.Tag = "graph";
            _mainGraph.Paint += MainGraph_Paint;
            graphLayout.Controls.Add(_mainGraph, 0, 1);

            TableLayoutPanel metrics = new TableLayoutPanel();
            metrics.Dock = DockStyle.Fill;
            metrics.ColumnCount = 3;
            metrics.RowCount = 2;
            metrics.Margin = new Padding(0, 8, 0, 0);
            metrics.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33.333f));
            metrics.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33.333f));
            metrics.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33.334f));
            metrics.RowStyles.Add(new RowStyle(SizeType.Percent, 50f));
            metrics.RowStyles.Add(new RowStyle(SizeType.Percent, 50f));
            _rootLayout.Controls.Add(metrics, 0, 2);

            GraphPanel dummy;
            metrics.Controls.Add(CreateMetricCard("Download", out _lblDownload, true, out _downloadGraph), 0, 0);
            metrics.Controls.Add(CreateMetricCard("Upload", out _lblUpload, true, out _uploadGraph), 1, 0);
            metrics.Controls.Add(CreateMetricCard("Latency", out _lblLatency, true, out _latencyGraph), 2, 0);
            metrics.Controls.Add(CreateMetricCard("Jitter", out _lblJitter, true, out _jitterGraph), 0, 1);
            metrics.Controls.Add(CreateMetricCard("Packet Loss", out _lblLoss, true, out _lossGraph), 1, 1);
            metrics.Controls.Add(CreateMetricCard("Interface", out _lblInterface, false, out dummy), 2, 1);

            _downloadGraph.Paint += delegate(object sender, PaintEventArgs e) { MiniGraph_Paint(sender as GraphPanel, e, MetricKind.Download); };
            _uploadGraph.Paint += delegate(object sender, PaintEventArgs e) { MiniGraph_Paint(sender as GraphPanel, e, MetricKind.Upload); };
            _latencyGraph.Paint += delegate(object sender, PaintEventArgs e) { MiniGraph_Paint(sender as GraphPanel, e, MetricKind.Latency); };
            _jitterGraph.Paint += delegate(object sender, PaintEventArgs e) { MiniGraph_Paint(sender as GraphPanel, e, MetricKind.Jitter); };
            _lossGraph.Paint += delegate(object sender, PaintEventArgs e) { MiniGraph_Paint(sender as GraphPanel, e, MetricKind.Loss); };

            Panel footer = CreateCard();
            footer.Dock = DockStyle.Fill;
            footer.Margin = new Padding(0);
            _rootLayout.Controls.Add(footer, 0, 3);

            Panel footerBody = new Panel();
            footerBody.Dock = DockStyle.Fill;
            footerBody.Padding = new Padding(10, 10, 10, 10);
            footer.Controls.Add(footerBody);

            _lblUpdated = new Label();
            _lblUpdated.AutoSize = false;
            _lblUpdated.AutoEllipsis = true;
            _lblUpdated.Tag = "muted";
            _lblUpdated.Dock = DockStyle.Fill;
            _lblUpdated.TextAlign = ContentAlignment.MiddleLeft;
            footerBody.Controls.Add(_lblUpdated);

            _btnClose = new Button();
            _btnClose.Text = "Close";
            _btnClose.Width = GetButtonWidth(_btnClose.Text, _btnClose.Font, 94);
            _btnClose.Height = Math.Max(30, _btnClose.Font.Height + 14);
            _btnClose.FlatStyle = FlatStyle.Flat;
            _btnClose.FlatAppearance.BorderSize = 1;
            _btnClose.AutoEllipsis = false;
            _btnClose.TextAlign = ContentAlignment.MiddleCenter;
            _btnClose.Dock = DockStyle.Right;
            _btnClose.Click += delegate { Close(); };
            footerBody.Controls.Add(_btnClose);

            ApplyResponsiveRowHeights();
        }

        private void InitializeGraphTooltips()
        {
            _graphToolTip = new ToolTip();
            _graphToolTip.InitialDelay = 120;
            _graphToolTip.ReshowDelay = 60;
            _graphToolTip.AutoPopDelay = 5000;
            _graphToolTip.ShowAlways = true;

            if (_mainGraph != null)
            {
                _mainGraph.MouseMove += MainGraph_MouseMove;
                _mainGraph.MouseLeave += Graph_MouseLeave;
            }

            BindMiniGraphTooltip(_downloadGraph, MetricKind.Download);
            BindMiniGraphTooltip(_uploadGraph, MetricKind.Upload);
            BindMiniGraphTooltip(_latencyGraph, MetricKind.Latency);
            BindMiniGraphTooltip(_jitterGraph, MetricKind.Jitter);
            BindMiniGraphTooltip(_lossGraph, MetricKind.Loss);
        }

        private void BindMiniGraphTooltip(GraphPanel panel, MetricKind kind)
        {
            if (panel == null)
            {
                return;
            }

            panel.MouseMove += delegate(object sender, MouseEventArgs e)
            {
                UpdateMiniGraphTooltip(panel, kind, e.Location);
            };
            panel.MouseLeave += Graph_MouseLeave;
        }

        private Panel CreateCard()
        {
            Panel panel = new Panel();
            panel.BackColor = _theme.CardBackground;
            panel.BorderStyle = BorderStyle.FixedSingle;
            panel.Tag = "card";
            panel.Margin = new Padding(0, 0, 0, 8);
            return panel;
        }

        private Panel CreateMetricCard(string title, out Label valueLabel, bool includeGraph, out GraphPanel graphPanel)
        {
            Panel card = CreateCard();
            card.Dock = DockStyle.Fill;
            card.Margin = new Padding(4);

            TableLayoutPanel layout = new TableLayoutPanel();
            layout.Dock = DockStyle.Fill;
            layout.Padding = new Padding(10, 8, 10, 8);
            layout.RowCount = includeGraph ? 3 : 2;
            layout.ColumnCount = 1;
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
            layout.RowStyles.Clear();
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            if (includeGraph)
            {
                layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));
            }
            card.Controls.Add(layout);

            Label titleLabel = new Label();
            titleLabel.AutoSize = true;
            titleLabel.Tag = "muted";
            titleLabel.Text = title;
            titleLabel.Margin = new Padding(0, 0, 0, 4);
            layout.Controls.Add(titleLabel, 0, 0);

            valueLabel = new Label();
            valueLabel.AutoSize = true;
            valueLabel.Font = new Font("Segoe UI Variable Text", 11f, FontStyle.Bold, GraphicsUnit.Point);
            valueLabel.AutoEllipsis = true;
            valueLabel.Margin = new Padding(0, 0, 0, includeGraph ? 4 : 0);
            layout.Controls.Add(valueLabel, 0, 1);

            graphPanel = null;
            if (includeGraph)
            {
                graphPanel = new GraphPanel();
                graphPanel.Dock = DockStyle.Fill;
                graphPanel.Tag = "spark";
                graphPanel.Margin = new Padding(0, 2, 0, 0);
                graphPanel.BorderStyle = BorderStyle.FixedSingle;
                layout.Controls.Add(graphPanel, 0, 2);
            }

            return card;
        }

        private void RenderTimer_Tick(object sender, EventArgs e)
        {
            if (!Visible || WindowState == FormWindowState.Minimized)
            {
                return;
            }

            GetTimeline(DateTime.UtcNow);
            InvalidateGraphs();
        }

        private void InvalidateGraphs()
        {
            if (_mainGraph != null) _mainGraph.Invalidate();
            if (_downloadGraph != null) _downloadGraph.Invalidate();
            if (_uploadGraph != null) _uploadGraph.Invalidate();
            if (_latencyGraph != null) _latencyGraph.Invalidate();
            if (_jitterGraph != null) _jitterGraph.Invalidate();
            if (_lossGraph != null) _lossGraph.Invalidate();
        }

        private void Graph_MouseLeave(object sender, EventArgs e)
        {
            Control control = sender as Control;
            if (control == null)
            {
                return;
            }
            HideGraphTooltip(control);
        }

        private void MainGraph_MouseMove(object sender, MouseEventArgs e)
        {
            if (_mainGraph == null)
            {
                return;
            }

            List<TimelineSample> samples = GetTimeline(DateTime.UtcNow);
            if (samples.Count == 0)
            {
                HideGraphTooltip(_mainGraph);
                return;
            }

            int bottomLabels;
            Rectangle plot = GetMainPlotRectangle(_mainGraph.ClientRectangle, out bottomLabels);
            if (e.Location.X < plot.Left || e.Location.X > plot.Right)
            {
                HideGraphTooltip(_mainGraph);
                return;
            }

            double xNorm = Clamp((double)(e.Location.X - plot.Left) / Math.Max(1, plot.Width), 0.0, 1.0);
            double offset = xNorm * WindowSeconds;
            int index = FindNearestSampleIndex(samples, offset, null);
            if (index < 0)
            {
                HideGraphTooltip(_mainGraph);
                return;
            }

            TimelineSample sample = samples[index];
            int secondsAgo = Math.Max(0, (int)Math.Round(WindowSeconds - sample.OffsetSec));
            string atText = secondsAgo == 0 ? "Now" : (secondsAgo.ToString() + "s ago");
            string tier = string.IsNullOrWhiteSpace(sample.Tier) ? "NoData" : sample.Tier;
            string scoreText = double.IsNaN(sample.QualityScore) ? "n/a" : sample.QualityScore.ToString("0") + "%";
            string text = string.Format("{0}  Score {1}  {2}", atText, scoreText, tier);
            ShowGraphTooltip(_mainGraph, text, e.Location);
        }

        private void UpdateMiniGraphTooltip(GraphPanel panel, MetricKind kind, Point location)
        {
            if (panel == null)
            {
                return;
            }

            Rectangle rect = panel.ClientRectangle;
            if (rect.Width < 2 || rect.Height < 2 || location.X < rect.Left || location.X > rect.Right)
            {
                HideGraphTooltip(panel);
                return;
            }

            List<TimelineSample> samples = GetTimeline(DateTime.UtcNow);
            if (samples.Count == 0)
            {
                HideGraphTooltip(panel);
                return;
            }

            double xNorm = Clamp((double)(location.X - rect.Left) / Math.Max(1, rect.Width), 0.0, 1.0);
            double offset = xNorm * WindowSeconds;
            int index = FindNearestSampleIndex(samples, offset, null);
            if (index < 0)
            {
                HideGraphTooltip(panel);
                return;
            }

            TimelineSample sample = samples[index];
            double value = GetMetric(sample, kind);
            int secondsAgo = Math.Max(0, (int)Math.Round(WindowSeconds - sample.OffsetSec));
            string atText = secondsAgo == 0 ? "Now" : (secondsAgo.ToString() + "s ago");
            string text = string.Format("{0}  {1} {2}", atText, GetMetricLabel(kind), FormatMetricValue(kind, value));
            ShowGraphTooltip(panel, text, location);
        }

        private void HideGraphTooltip(Control control)
        {
            if (_graphToolTip == null || control == null)
            {
                return;
            }

            if (_activeTooltipControl == control)
            {
                _graphToolTip.Hide(control);
                _activeTooltipControl = null;
                _activeTooltipText = string.Empty;
            }
        }

        private void ShowGraphTooltip(Control control, string text, Point location)
        {
            if (_graphToolTip == null || control == null)
            {
                return;
            }

            string safe = text ?? string.Empty;
            if (safe.Length == 0)
            {
                HideGraphTooltip(control);
                return;
            }

            bool changedControl = _activeTooltipControl != control;
            bool changedText = !string.Equals(_activeTooltipText, safe, StringComparison.Ordinal);
            _graphToolTip.Show(safe, control, location.X + 14, location.Y + 16, 1100);
            _activeTooltipControl = control;
            if (changedControl || changedText)
            {
                _activeTooltipText = safe;
            }
        }

        private void MainGraph_Paint(object sender, PaintEventArgs e)
        {
            GraphPanel panel = sender as GraphPanel;
            if (panel == null) return;

            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
            Rectangle rect = panel.ClientRectangle;
            if (rect.Width < 100 || rect.Height < 70) return;

            int bottomLabels;
            Rectangle plot = GetMainPlotRectangle(rect, out bottomLabels);
            if (plot.Width < 30 || plot.Height < 30) return;

            Color bg = _theme.IsDark ? Color.FromArgb(31, 41, 55) : Color.FromArgb(246, 250, 255);
            Color border = _theme.IsDark ? Color.FromArgb(66, 78, 94) : Color.FromArgb(204, 216, 228);
            using (SolidBrush b = new SolidBrush(bg)) { e.Graphics.FillRectangle(b, plot); }
            using (Pen p = new Pen(border, 1f)) { e.Graphics.DrawRectangle(p, plot); }

            List<TimelineSample> samples = GetTimeline(DateTime.UtcNow);
            if (samples.Count < 2 || !HasQualityData(samples))
            {
                TextRenderer.DrawText(e.Graphics, "Collecting quality history...", Font, plot, _theme.TextSecondary, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
                DrawMainAxisLabels(e.Graphics, plot, bottomLabels);
                return;
            }

            List<PointF> points = new List<PointF>(samples.Count);
            List<bool> valid = new List<bool>(samples.Count);
            for (int i = 0; i < samples.Count; i++)
            {
                float x = plot.Left + (float)((samples[i].OffsetSec / WindowSeconds) * plot.Width);
                bool hasData = !double.IsNaN(samples[i].QualityScore);
                float y = hasData
                    ? plot.Bottom - (float)(Clamp(samples[i].QualityScore / 100.0, 0.0, 1.0) * plot.Height)
                    : plot.Bottom;
                points.Add(new PointF(x, y));
                valid.Add(hasData);
            }

            int iRun = 1;
            while (iRun < points.Count)
            {
                if (!valid[iRun - 1] || !valid[iRun])
                {
                    iRun++;
                    continue;
                }

                string tier = NormalizeTier(samples[iRun - 1].Tier);
                int start = iRun - 1;
                int end = iRun;
                iRun++;
                while (iRun < points.Count
                    && valid[iRun - 1]
                    && valid[iRun]
                    && string.Equals(NormalizeTier(samples[iRun - 1].Tier), tier, StringComparison.Ordinal))
                {
                    end = iRun;
                    iRun++;
                }

                DrawMainGraphRun(e.Graphics, points, start, end, plot.Bottom, GetTierColor(tier));
            }

            int lastValidIndex = -1;
            for (int i = valid.Count - 1; i >= 0; i--)
            {
                if (valid[i])
                {
                    lastValidIndex = i;
                    break;
                }
            }

            if (lastValidIndex >= 0)
            {
                PointF last = points[lastValidIndex];
                Color lastColor = GetTierColor(samples[lastValidIndex].Tier);
                Rectangle marker = new Rectangle((int)last.X - 4, (int)last.Y - 4, 8, 8);
                using (SolidBrush sb = new SolidBrush(lastColor)) { e.Graphics.FillEllipse(sb, marker); }
            }

            DrawMainAxisLabels(e.Graphics, plot, bottomLabels);
        }

        private Rectangle GetMainPlotRectangle(Rectangle panelRect, out int bottomLabels)
        {
            int margin = 8;
            bottomLabels = 30;
            return new Rectangle(
                panelRect.Left + margin,
                panelRect.Top + margin,
                panelRect.Width - (margin * 2),
                panelRect.Height - (margin * 2) - bottomLabels);
        }

        private void DrawMainAxisLabels(Graphics graphics, Rectangle plot, int bottomLabels)
        {
            int labelY = plot.Bottom + 4;
            TextRenderer.DrawText(graphics, "60s", Font, new Rectangle(plot.Left, labelY, 72, bottomLabels), _theme.TextPrimary, TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPadding);
            TextRenderer.DrawText(graphics, "Now", Font, new Rectangle(plot.Right - 72, labelY, 72, bottomLabels), _theme.TextPrimary, TextFormatFlags.Right | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPadding);
        }

        private void DrawMainGraphRun(Graphics graphics, List<PointF> points, int start, int end, float baselineY, Color color)
        {
            if (start < 0 || end <= start || end >= points.Count)
            {
                return;
            }

            List<PointF> runPoints = new List<PointF>(end - start + 1);
            for (int i = start; i <= end; i++)
            {
                runPoints.Add(points[i]);
            }

            if (runPoints.Count < 2)
            {
                return;
            }

            List<PointF> fillPoly = new List<PointF>(runPoints.Count + 2);
            fillPoly.Add(new PointF(runPoints[0].X, baselineY));
            fillPoly.AddRange(runPoints);
            fillPoly.Add(new PointF(runPoints[runPoints.Count - 1].X, baselineY));

            using (SolidBrush areaBrush = new SolidBrush(Color.FromArgb(_theme.IsDark ? 52 : 40, color)))
            {
                graphics.FillPolygon(areaBrush, fillPoly.ToArray());
            }

            Color glow = _theme.IsDark
                ? Color.FromArgb(80, Math.Min(255, color.R + 28), Math.Min(255, color.G + 28), Math.Min(255, color.B + 28))
                : Color.FromArgb(58, color);
            using (Pen glowPen = new Pen(glow, 3.6f))
            using (Pen linePen = new Pen(color, 2.2f))
            {
                glowPen.StartCap = LineCap.Round;
                glowPen.EndCap = LineCap.Round;
                linePen.StartCap = LineCap.Round;
                linePen.EndCap = LineCap.Round;

                PointF[] runArray = runPoints.ToArray();
                graphics.DrawLines(glowPen, runArray);
                graphics.DrawLines(linePen, runArray);
            }
        }

        private void MiniGraph_Paint(GraphPanel panel, PaintEventArgs e, MetricKind kind)
        {
            if (panel == null || e == null) return;

            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
            Rectangle rect = panel.ClientRectangle;
            if (rect.Width < 30 || rect.Height < 18) return;

            Color sparkBg = _theme.IsDark ? Color.FromArgb(34, 43, 56) : Color.FromArgb(244, 248, 252);
            using (SolidBrush sb = new SolidBrush(sparkBg)) { e.Graphics.FillRectangle(sb, rect); }

            List<TimelineSample> samples = GetTimeline(DateTime.UtcNow);
            if (samples.Count < 2) return;

            double metricMin;
            double metricMax;
            bool lowerIsBetter;
            GetMetricScale(kind, samples, out metricMin, out metricMax, out lowerIsBetter);

            List<PointF> points = new List<PointF>(samples.Count);
            List<bool> valid = new List<bool>(samples.Count);
            for (int i = 0; i < samples.Count; i++)
            {
                double value = GetMetric(samples[i], kind);
                double xNorm = samples[i].OffsetSec / WindowSeconds;
                bool hasData = !double.IsNaN(value);
                double yNorm = 0.0;
                if (hasData)
                {
                    if (kind == MetricKind.Download || kind == MetricKind.Upload)
                    {
                        double denom = Math.Log10(metricMax + 1.0);
                        yNorm = denom <= 0.0 ? 0.0 : Clamp(Math.Log10(Math.Max(0.0, value) + 1.0) / denom, 0.0, 1.0);
                    }
                    else
                    {
                        yNorm = Clamp((value - metricMin) / (metricMax - metricMin), 0.0, 1.0);
                    }
                    if (lowerIsBetter)
                    {
                        yNorm = 1.0 - yNorm;
                    }
                }

                float x = rect.Left + (float)(xNorm * rect.Width);
                float y = hasData
                    ? rect.Top + (float)((1.0 - yNorm) * (rect.Height - 1))
                    : rect.Bottom;
                points.Add(new PointF(x, y));
                valid.Add(hasData);
            }

            Color lineColor = GetMetricColor(kind);
            int index = 1;
            while (index < points.Count)
            {
                if (!valid[index - 1] || !valid[index])
                {
                    index++;
                    continue;
                }

                int runStart = index - 1;
                int runEnd = index;
                index++;
                while (index < points.Count && valid[index - 1] && valid[index])
                {
                    runEnd = index;
                    index++;
                }

                DrawMiniGraphRun(e.Graphics, points, runStart, runEnd, rect.Bottom, lineColor);
            }
        }

        private static void GetMetricScale(MetricKind kind, List<TimelineSample> samples, out double min, out double max, out bool lowerIsBetter)
        {
            min = 0.0;
            lowerIsBetter = false;
            double observedMax = GetObservedMetricMax(samples, kind);
            switch (kind)
            {
                case MetricKind.Download:
                    max = Math.Max(40.0, RoundUpNice(Math.Max(observedMax, 20.0) * 1.15, 25.0));
                    break;
                case MetricKind.Upload:
                    max = Math.Max(20.0, RoundUpNice(Math.Max(observedMax, 10.0) * 1.15, 10.0));
                    break;
                case MetricKind.Latency:
                    max = Math.Max(80.0, RoundUpNice(Math.Max(observedMax, 40.0) * 1.20, 20.0));
                    lowerIsBetter = true;
                    break;
                case MetricKind.Jitter:
                    max = Math.Max(20.0, RoundUpNice(Math.Max(observedMax, 10.0) * 1.25, 10.0));
                    break;
                case MetricKind.Loss:
                    max = 100.0;
                    break;
                default:
                    max = 100.0;
                    break;
            }

            if (max <= min)
            {
                max = min + 1.0;
            }
        }

        private static double GetObservedMetricMax(List<TimelineSample> samples, MetricKind kind)
        {
            if (samples == null || samples.Count == 0)
            {
                return 0.0;
            }

            double max = 0.0;
            for (int i = 0; i < samples.Count; i++)
            {
                double value = GetMetric(samples[i], kind);
                if (!double.IsNaN(value) && !double.IsInfinity(value) && value > max)
                {
                    max = value;
                }
            }

            return max;
        }

        private static double RoundUpNice(double value, double step)
        {
            if (step <= 0.0)
            {
                return value;
            }

            return Math.Ceiling(Math.Max(0.0, value) / step) * step;
        }

        private void DrawMiniGraphRun(Graphics graphics, List<PointF> points, int start, int end, float baselineY, Color lineColor)
        {
            if (start < 0 || end <= start || end >= points.Count)
            {
                return;
            }

            List<PointF> runPoints = new List<PointF>(end - start + 1);
            for (int i = start; i <= end; i++)
            {
                runPoints.Add(points[i]);
            }

            if (runPoints.Count < 2)
            {
                return;
            }

            List<PointF> area = new List<PointF>(runPoints.Count + 2);
            area.Add(new PointF(runPoints[0].X, baselineY));
            area.AddRange(runPoints);
            area.Add(new PointF(runPoints[runPoints.Count - 1].X, baselineY));
            using (SolidBrush fill = new SolidBrush(Color.FromArgb(_theme.IsDark ? 30 : 26, lineColor)))
            {
                graphics.FillPolygon(fill, area.ToArray());
            }

            using (Pen pen = new Pen(lineColor, 1.9f))
            {
                pen.StartCap = LineCap.Round;
                pen.EndCap = LineCap.Round;
                graphics.DrawLines(pen, runPoints.ToArray());
            }
        }

        private void InvalidateTimelineCache()
        {
            _timelineCacheBuiltAtUtc = DateTime.MinValue;
            _timelineCacheHistoryVersion = -1;
            if (_timelineCache != null)
            {
                _timelineCache.Clear();
            }
        }

        private List<TimelineSample> GetTimeline(DateTime nowUtc)
        {
            if (_timelineCache != null
                && _timelineCacheHistoryVersion == _historyVersion
                && Math.Abs((nowUtc - _timelineCacheBuiltAtUtc).TotalMilliseconds) <= TimelineCacheMaxAgeMs
                && _timelineCache.Count > 0)
            {
                return _timelineCache;
            }

            _timelineCache = BuildTimelineCore(nowUtc, TimelineStepSeconds);
            _timelineCacheBuiltAtUtc = nowUtc;
            _timelineCacheHistoryVersion = _historyVersion;
            return _timelineCache;
        }

        private List<TimelineSample> BuildTimelineCore(DateTime nowUtc, double stepSec)
        {
            List<TimelineSample> output = new List<TimelineSample>();
            if (_history == null || _history.Count == 0)
            {
                return output;
            }

            List<QualityHistoryPoint> ordered = new List<QualityHistoryPoint>(_history.Count);
            bool outOfOrder = false;
            DateTime previous = DateTime.MinValue;
            for (int i = 0; i < _history.Count; i++)
            {
                QualityHistoryPoint point = _history[i];
                if (point == null)
                {
                    continue;
                }

                if (ordered.Count > 0 && point.TimestampUtc < previous)
                {
                    outOfOrder = true;
                }

                ordered.Add(point);
                previous = point.TimestampUtc;
            }

            if (ordered.Count == 0)
            {
                return output;
            }

            if (outOfOrder)
            {
                ordered.Sort(delegate(QualityHistoryPoint left, QualityHistoryPoint right)
                {
                    return left.TimestampUtc.CompareTo(right.TimestampUtc);
                });
            }

            DateTime start = nowUtc.AddSeconds(-WindowSeconds);
            int idx = 0;
            for (double offset = 0.0; offset <= WindowSeconds; offset += stepSec)
            {
                DateTime t = start.AddSeconds(offset);
                if (t < ordered[0].TimestampUtc)
                {
                    output.Add(CreateEmptySample(offset));
                    continue;
                }

                while (idx + 1 < ordered.Count && ordered[idx + 1].TimestampUtc <= t)
                {
                    idx++;
                }

                QualityHistoryPoint left = ordered[Math.Min(idx, ordered.Count - 1)];
                QualityHistoryPoint right = (idx + 1 < ordered.Count) ? ordered[idx + 1] : left;

                double interp = 0.0;
                if (right.TimestampUtc > left.TimestampUtc)
                {
                    interp = (t - left.TimestampUtc).TotalMilliseconds / (right.TimestampUtc - left.TimestampUtc).TotalMilliseconds;
                    interp = Clamp(interp, 0.0, 1.0);
                }

                double eased = Ease(interp);
                TimelineSample sample = new TimelineSample();
                sample.OffsetSec = offset;
                sample.QualityScore = SmoothScore(ordered, idx, eased);
                sample.Tier = !string.IsNullOrWhiteSpace(left.Tier)
                    ? left.Tier
                    : (string.IsNullOrWhiteSpace(right.Tier) ? _currentTier : right.Tier);
                sample.DownloadMbps = Lerp(left.DownloadMbps, right.DownloadMbps, eased);
                sample.UploadMbps = Lerp(left.UploadMbps, right.UploadMbps, eased);
                sample.LatencyMs = Lerp(left.LatencyMs, right.LatencyMs, eased);
                sample.JitterMs = Lerp(left.JitterMs, right.JitterMs, eased);
                sample.LossPct = Lerp(left.LossPct, right.LossPct, eased);
                output.Add(sample);
            }

            return output;
        }

        private static TimelineSample CreateEmptySample(double offset)
        {
            TimelineSample empty = new TimelineSample();
            empty.OffsetSec = offset;
            empty.QualityScore = double.NaN;
            empty.Tier = null;
            empty.DownloadMbps = double.NaN;
            empty.UploadMbps = double.NaN;
            empty.LatencyMs = double.NaN;
            empty.JitterMs = double.NaN;
            empty.LossPct = double.NaN;
            return empty;
        }

        private static double Lerp(double a, double b, double t)
        {
            if (double.IsNaN(a) && double.IsNaN(b)) return double.NaN;
            if (double.IsNaN(a)) return b;
            if (double.IsNaN(b)) return a;
            return a + ((b - a) * t);
        }

        private static double SmoothScore(List<QualityHistoryPoint> ordered, int index, double interp)
        {
            int p0 = Math.Max(0, index - 1);
            int p1 = Math.Max(0, Math.Min(index, ordered.Count - 1));
            int p2 = Math.Max(0, Math.Min(index + 1, ordered.Count - 1));
            int p3 = Math.Max(0, Math.Min(index + 2, ordered.Count - 1));

            double v0 = ordered[p0].QualityScore;
            double v1 = ordered[p1].QualityScore;
            double v2 = ordered[p2].QualityScore;
            double v3 = ordered[p3].QualityScore;

            double t = Clamp(interp, 0.0, 1.0);
            double t2 = t * t;
            double t3 = t2 * t;

            double value = 0.5 * (
                (2.0 * v1) +
                (-v0 + v2) * t +
                ((2.0 * v0) - (5.0 * v1) + (4.0 * v2) - v3) * t2 +
                (-v0 + (3.0 * v1) - (3.0 * v2) + v3) * t3
            );
            return Clamp(value, 0.0, 100.0);
        }

        private static double Ease(double value)
        {
            double t = Clamp(value, 0.0, 1.0);
            return t * t * (3.0 - (2.0 * t));
        }

        private static int FindNearestSampleIndex(List<TimelineSample> samples, double offsetSec, Func<TimelineSample, bool> predicate)
        {
            int bestIndex = -1;
            double bestDistance = double.MaxValue;
            for (int i = 0; i < samples.Count; i++)
            {
                TimelineSample sample = samples[i];
                if (predicate != null && !predicate(sample))
                {
                    continue;
                }

                double distance = Math.Abs(sample.OffsetSec - offsetSec);
                if (distance < bestDistance)
                {
                    bestDistance = distance;
                    bestIndex = i;
                }
            }

            return bestIndex;
        }

        private static string NormalizeTier(string tier)
        {
            string text = (tier ?? string.Empty).Trim();
            return text.Length == 0 ? "Offline" : text;
        }

        private static string GetMetricLabel(MetricKind kind)
        {
            switch (kind)
            {
                case MetricKind.Download: return "Download";
                case MetricKind.Upload: return "Upload";
                case MetricKind.Latency: return "Latency";
                case MetricKind.Jitter: return "Jitter";
                case MetricKind.Loss: return "Packet Loss";
                default: return "Value";
            }
        }

        private static Color GetMetricColor(MetricKind kind)
        {
            switch (kind)
            {
                case MetricKind.Download: return Color.FromArgb(56, 189, 248);
                case MetricKind.Upload: return Color.FromArgb(34, 197, 94);
                case MetricKind.Latency: return Color.FromArgb(251, 191, 36);
                case MetricKind.Jitter: return Color.FromArgb(244, 114, 182);
                case MetricKind.Loss: return Color.FromArgb(248, 113, 113);
                default: return Color.FromArgb(90, 179, 255);
            }
        }

        private static string FormatMetricValue(MetricKind kind, double value)
        {
            if (double.IsNaN(value))
            {
                return "n/a";
            }

            switch (kind)
            {
                case MetricKind.Download:
                case MetricKind.Upload:
                    return value.ToString("0.00") + " Mbps";
                case MetricKind.Latency:
                case MetricKind.Jitter:
                    return value.ToString("0.0") + " ms";
                case MetricKind.Loss:
                    return value.ToString("0.0") + "%";
                default:
                    return value.ToString("0.0");
            }
        }

        private static double GetMetric(TimelineSample sample, MetricKind kind)
        {
            switch (kind)
            {
                case MetricKind.Download: return sample.DownloadMbps;
                case MetricKind.Upload: return sample.UploadMbps;
                case MetricKind.Latency: return sample.LatencyMs;
                case MetricKind.Jitter: return sample.JitterMs;
                case MetricKind.Loss: return sample.LossPct;
                default: return double.NaN;
            }
        }

        private static bool HasQualityData(List<TimelineSample> samples)
        {
            if (samples == null || samples.Count == 0)
            {
                return false;
            }

            for (int i = 0; i < samples.Count; i++)
            {
                if (!double.IsNaN(samples[i].QualityScore))
                {
                    return true;
                }
            }

            return false;
        }

        private static int GetButtonWidth(string text, Font font, int minimumWidth)
        {
            string safe = text ?? string.Empty;
            Font effective = font ?? SystemFonts.MessageBoxFont;
            Size size = TextRenderer.MeasureText(
                safe,
                effective,
                new Size(int.MaxValue, int.MaxValue),
                TextFormatFlags.SingleLine | TextFormatFlags.NoPrefix | TextFormatFlags.NoPadding);
            int width = size.Width + 24;
            return Math.Max(minimumWidth, width);
        }

        private static double Clamp(double value, double min, double max)
        {
            if (value < min) return min;
            if (value > max) return max;
            return value;
        }

        private void StatusForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            SystemEvents.UserPreferenceChanged -= SystemEvents_UserPreferenceChanged;
            if (_renderTimer != null)
            {
                try { _renderTimer.Stop(); } catch { }
                try { _renderTimer.Tick -= RenderTimer_Tick; } catch { }
                try { _renderTimer.Dispose(); } catch { }
                _renderTimer = null;
            }

            if (_graphToolTip != null)
            {
                try { _graphToolTip.RemoveAll(); } catch { }
                try { _graphToolTip.Dispose(); } catch { }
                _graphToolTip = null;
            }
            _activeTooltipControl = null;
            _activeTooltipText = string.Empty;
        }

        private void SystemEvents_UserPreferenceChanged(object sender, UserPreferenceChangedEventArgs e)
        {
            if (IsDisposed) return;
            try
            {
                BeginInvoke((Action)(() =>
                {
                    if (!IsDisposed) ApplyTheme();
                }));
            }
            catch { }
        }

        private void ApplyTheme()
        {
            _theme = UiTheme.GetCurrent();
            BackColor = _theme.AppBackground;
            ForeColor = _theme.TextPrimary;
            NativeUi.ApplyWindowTheme(this, _theme.IsDark);

            ApplyThemeRecursive(this);
            if (_btnClose != null)
            {
                _btnClose.Width = GetButtonWidth(_btnClose.Text, _btnClose.Font, 94);
                _btnClose.BackColor = _theme.SecondaryButtonBackground;
                _btnClose.ForeColor = _theme.SecondaryButtonText;
                _btnClose.FlatAppearance.BorderColor = _theme.SecondaryButtonBorder;
            }

            InvalidateGraphs();
        }

        private void ApplyThemeRecursive(Control c)
        {
            if (c == null) return;

            if (c is Panel)
            {
                if (Equals(c.Tag, "card")) c.BackColor = _theme.CardBackground;
                else if (Equals(c.Tag, "graph") || Equals(c.Tag, "spark")) c.BackColor = _theme.IsDark ? Color.FromArgb(30, 38, 48) : Color.FromArgb(246, 250, 255);
                else c.BackColor = _theme.AppBackground;
                c.ForeColor = _theme.TextPrimary;
            }
            else if (c is Label)
            {
                c.BackColor = Color.Transparent;
                c.ForeColor = Equals(c.Tag, "muted") ? _theme.TextSecondary : _theme.TextPrimary;
            }
            else if (c is TableLayoutPanel || c is FlowLayoutPanel)
            {
                bool insideCard = c.Parent != null && Equals(c.Parent.Tag, "card");
                c.BackColor = insideCard ? _theme.CardBackground : _theme.AppBackground;
                c.ForeColor = _theme.TextPrimary;
            }

            foreach (Control child in c.Controls) ApplyThemeRecursive(child);
        }

        private static Color GetTierColor(string tier)
        {
            switch ((tier ?? string.Empty).Trim())
            {
                case "High": return Color.FromArgb(46, 204, 113);
                case "Poor": return Color.FromArgb(241, 196, 15);
                case "VeryPoor": return Color.FromArgb(230, 126, 34);
                case "Bad": return Color.FromArgb(231, 76, 60);
                case "Offline": return Color.FromArgb(149, 17, 17);
                case "Paused": return Color.FromArgb(160, 160, 160);
                default: return Color.FromArgb(150, 156, 166);
            }
        }

        private enum MetricKind { Download, Upload, Latency, Jitter, Loss }

        private sealed class TimelineSample
        {
            public double OffsetSec { get; set; }
            public double QualityScore { get; set; }
            public string Tier { get; set; }
            public double DownloadMbps { get; set; }
            public double UploadMbps { get; set; }
            public double LatencyMs { get; set; }
            public double JitterMs { get; set; }
            public double LossPct { get; set; }
        }

        private sealed class GraphPanel : Panel
        {
            public GraphPanel()
            {
                DoubleBuffered = true;
                ResizeRedraw = true;
            }
        }
    }
}
