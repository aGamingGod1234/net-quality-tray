using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Microsoft.Win32;

namespace NetQualitySentinel
{
    internal sealed class SettingsForm : Form
    {
        private readonly string _appName;
        private readonly AppConfig _workingConfig;
        private readonly Snapshot _snapshot;
        private readonly bool _initialPaused;
        private readonly bool _initialStartup;
        private readonly Dictionary<string, TextBox> _colorBoxes;
        private readonly Dictionary<string, Panel> _colorSwatches;
        private readonly ColorDialog _colorDialog;
        private UiTheme _theme;

        private CheckBox _chkStartup;
        private CheckBox _chkPaused;
        private Label _lblLive;
        private Button _btnStatusDetails;
        private Label _lblRanges;
        private TrackBar _trkHigh;
        private TrackBar _trkPoor;
        private TrackBar _trkVeryPoor;
        private NumericUpDown _numHigh;
        private NumericUpDown _numPoor;
        private NumericUpDown _numVeryPoor;
        private ComboBox _cmbTargetTier;
        private bool _syncingThresholds;
        private SplitContainer _bodySplit;
        private FlowLayoutPanel _leftColumn;
        private FlowLayoutPanel _rightColumn;
        private Panel _thresholdCard;
        private Panel _colorCard;
        private FlowLayoutPanel _chipPanel;
        private TableLayoutPanel _colorGrid;
        private TableLayoutPanel _footerTable;
        private FlowLayoutPanel _footerActions;
        private Label _footerText;
        private TableLayoutPanel _rootTable;
        private Panel _headerCard;
        private Label _headerTitle;
        private Label _headerSubtitle;

        public bool Saved { get; private set; }
        public bool ForceProbe { get; private set; }
        public bool ExitRequested { get; private set; }
        public bool StartupEnabled { get; private set; }
        public bool Paused { get; private set; }
        public AppConfig UpdatedConfig { get; private set; }

        public SettingsForm(
            string appName,
            AppConfig config,
            Snapshot snapshot,
            bool paused,
            bool startupEnabled,
            string iconPath)
        {
            _appName = appName;
            _workingConfig = config.Clone();
            _snapshot = snapshot;
            _initialPaused = paused;
            _initialStartup = startupEnabled;
            _colorBoxes = new Dictionary<string, TextBox>(StringComparer.OrdinalIgnoreCase);
            _colorSwatches = new Dictionary<string, Panel>(StringComparer.OrdinalIgnoreCase);
            _colorDialog = new ColorDialog();

            BuildWindow(iconPath);
            BuildLayout();
            BindInitialValues();
            ApplyTheme();
            ApplyResponsiveLayout();
            SystemEvents.UserPreferenceChanged += SystemEvents_UserPreferenceChanged;
            FormClosed += SettingsForm_FormClosed;
        }

        private void BuildWindow(string iconPath)
        {
            _theme = UiTheme.GetCurrent();
            Text = _appName + " Settings";
            StartPosition = FormStartPosition.CenterScreen;
            MinimumSize = new Size(920, 680);
            Size = new Size(1120, 760);
            BackColor = _theme.AppBackground;
            ForeColor = _theme.TextPrimary;
            Font = new Font("Segoe UI", 10f, FontStyle.Regular, GraphicsUnit.Point);
            AutoScaleMode = AutoScaleMode.Dpi;
            FormBorderStyle = FormBorderStyle.Sizable;
            MaximizeBox = true;
            MinimizeBox = true;
            ShowInTaskbar = true;
            HandleCreated += delegate { NativeUi.ApplyWindowTheme(this, _theme.IsDark); };

            if (File.Exists(iconPath))
            {
                try
                {
                    Icon = new Icon(iconPath);
                }
                catch
                {
                }
            }
        }

        private void BuildLayout()
        {
            _rootTable = new TableLayoutPanel();
            _rootTable.Dock = DockStyle.Fill;
            _rootTable.ColumnCount = 1;
            _rootTable.RowCount = 3;
            _rootTable.Padding = new Padding(12);
            _rootTable.RowStyles.Add(new RowStyle(SizeType.Absolute, 116));
            _rootTable.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            _rootTable.RowStyles.Add(new RowStyle(SizeType.Absolute, 124));
            Controls.Add(_rootTable);

            _headerCard = CreateCardPanel();
            _headerCard.Padding = new Padding(12);
            _headerCard.Margin = new Padding(0, 0, 0, 8);
            _rootTable.Controls.Add(_headerCard, 0, 0);

            PictureBox logo = new PictureBox();
            logo.Size = new Size(48, 48);
            logo.Location = new Point(12, 12);
            logo.BackColor = Color.Transparent;
            logo.Image = BuildLogoBitmap(48, 48);
            logo.SizeMode = PictureBoxSizeMode.CenterImage;
            _headerCard.Controls.Add(logo);

            _headerTitle = new Label();
            _headerTitle.AutoSize = true;
            _headerTitle.Font = new Font("Segoe UI Variable Text", 15f, FontStyle.Bold, GraphicsUnit.Point);
            _headerTitle.Text = _appName;
            _headerTitle.Location = new Point(74, 16);
            _headerCard.Controls.Add(_headerTitle);

            _headerSubtitle = new Label();
            _headerSubtitle.AutoSize = false;
            _headerSubtitle.AutoEllipsis = true;
            _headerSubtitle.Font = new Font("Segoe UI", 9f, FontStyle.Regular, GraphicsUnit.Point);
            _headerSubtitle.ForeColor = _theme.TextSecondary;
            _headerSubtitle.Tag = "muted";
            _headerSubtitle.Text = "Clean Windows-style controls with immediate settings apply";
            _headerSubtitle.Location = new Point(86, 52);
            _headerSubtitle.Size = new Size(740, 24);
            _headerSubtitle.Visible = false;
            _headerCard.Controls.Add(_headerSubtitle);

            _bodySplit = new SplitContainer();
            _bodySplit.Dock = DockStyle.Fill;
            _bodySplit.FixedPanel = FixedPanel.Panel1;
            _bodySplit.Panel1MinSize = 0;
            _bodySplit.Panel2MinSize = 0;
            _bodySplit.IsSplitterFixed = false;
            _rootTable.Controls.Add(_bodySplit, 0, 1);

            _leftColumn = new FlowLayoutPanel();
            _leftColumn.Dock = DockStyle.Fill;
            _leftColumn.FlowDirection = FlowDirection.TopDown;
            _leftColumn.WrapContents = false;
            _leftColumn.AutoScroll = true;
            _leftColumn.Padding = new Padding(0);
            _bodySplit.Panel1.Controls.Add(_leftColumn);

            _rightColumn = new FlowLayoutPanel();
            _rightColumn.Dock = DockStyle.Fill;
            _rightColumn.FlowDirection = FlowDirection.TopDown;
            _rightColumn.WrapContents = false;
            _rightColumn.AutoScroll = true;
            _rightColumn.Padding = new Padding(0, 0, 6, 0);
            _bodySplit.Panel2.Controls.Add(_rightColumn);

            Panel liveCard = CreateCardPanel();
            liveCard.Name = "LiveCard";
            liveCard.Size = new Size(312, 176);
            liveCard.Margin = new Padding(0, 0, 0, 10);
            _leftColumn.Controls.Add(liveCard);
            AddCardHeading(liveCard, "Live Status", "Current network quality snapshot");
            _lblLive = new Label();
            _lblLive.Name = "LiveBodyLabel";
            _lblLive.AutoSize = false;
            _lblLive.Location = new Point(14, 96);
            _lblLive.Size = new Size(284, 70);
            _lblLive.ForeColor = _theme.TextPrimary;
            _lblLive.Text = "Waiting for first probe...";
            liveCard.Controls.Add(_lblLive);

            Panel generalCard = CreateCardPanel();
            generalCard.Name = "GeneralCard";
            generalCard.Size = new Size(312, 176);
            generalCard.Margin = new Padding(0, 0, 0, 10);
            _leftColumn.Controls.Add(generalCard);
            AddCardHeading(generalCard, "General", "Core runtime preferences");
            _chkStartup = new CheckBox();
            _chkStartup.Text = "Start with Windows";
            _chkStartup.AutoSize = true;
            _chkStartup.Location = new Point(16, 96);
            generalCard.Controls.Add(_chkStartup);
            _chkPaused = new CheckBox();
            _chkPaused.Text = "Pause monitoring";
            _chkPaused.AutoSize = true;
            _chkPaused.Location = new Point(16, 126);
            generalCard.Controls.Add(_chkPaused);

            Panel actionsCard = CreateCardPanel();
            actionsCard.Name = "ActionsCard";
            actionsCard.Size = new Size(312, 146);
            actionsCard.Margin = new Padding(0, 0, 0, 10);
            _leftColumn.Controls.Add(actionsCard);
            AddCardHeading(actionsCard, "Actions", "Quick runtime actions");

            _btnStatusDetails = CreateSecondaryButton("Status Details", 124);
            _btnStatusDetails.Location = new Point(16, 96);
            _btnStatusDetails.Click += delegate
            {
                ShowSnapshotDialog();
            };
            actionsCard.Controls.Add(_btnStatusDetails);

            _thresholdCard = CreateCardPanel();
            _thresholdCard.Size = new Size(690, 340);
            _thresholdCard.Margin = new Padding(0, 0, 0, 10);
            _rightColumn.Controls.Add(_thresholdCard);
            AddCardHeading(_thresholdCard, "Quality Thresholds", "Define the score boundaries for each quality tier");

            AddThresholdRow(_thresholdCard, "High starts at", 116, out _trkHigh, out _numHigh, 1, 100);
            AddThresholdRow(_thresholdCard, "Poor starts at", 170, out _trkPoor, out _numPoor, 0, 99);
            AddThresholdRow(_thresholdCard, "Very Poor starts at", 224, out _trkVeryPoor, out _numVeryPoor, 0, 98);
            _lblRanges = new Label();
            _lblRanges.AutoSize = false;
            _lblRanges.Location = new Point(18, 286);
            _lblRanges.Size = new Size(650, 40);
            _lblRanges.ForeColor = _theme.TextSecondary;
            _lblRanges.Tag = "muted";
            _thresholdCard.Controls.Add(_lblRanges);

            _colorCard = CreateCardPanel();
            _colorCard.Size = new Size(690, 470);
            _colorCard.Margin = new Padding(0, 0, 0, 10);
            _rightColumn.Controls.Add(_colorCard);
            AddCardHeading(_colorCard, "Quality Colors", "Apply palette chips or edit hex values per tier");

            Label applyLabel = new Label();
            applyLabel.Text = "Apply palette to:";
            applyLabel.AutoSize = true;
            applyLabel.ForeColor = _theme.TextSecondary;
            applyLabel.Tag = "muted";
            applyLabel.Location = new Point(16, 96);
            _colorCard.Controls.Add(applyLabel);

            _cmbTargetTier = new ComboBox();
            _cmbTargetTier.DropDownStyle = ComboBoxStyle.DropDownList;
            _cmbTargetTier.Items.AddRange(new object[] { "High", "Poor", "VeryPoor", "Bad", "Offline", "Paused" });
            _cmbTargetTier.SelectedIndex = 0;
            _cmbTargetTier.Location = new Point(168, 90);
            _cmbTargetTier.Width = 160;
            _colorCard.Controls.Add(_cmbTargetTier);

            _chipPanel = new FlowLayoutPanel();
            _chipPanel.Location = new Point(16, 136);
            _chipPanel.Size = new Size(648, 80);
            _chipPanel.WrapContents = true;
            _chipPanel.BackColor = _theme.ChipSurface;
            _chipPanel.Padding = new Padding(8, 8, 8, 6);
            _chipPanel.BorderStyle = BorderStyle.FixedSingle;
            _colorCard.Controls.Add(_chipPanel);

            string[] chips = new string[]
            {
                "#2ECC71", "#27AE60", "#00A37A", "#0EA5E9", "#4CAF50", "#8BC34A",
                "#F1C40F", "#FFC107", "#FF9800", "#FDBA74", "#FFE082", "#FFD54F",
                "#E67E22", "#FB8C00", "#FF7043", "#FF5722", "#F4511E", "#EF6C00",
                "#E74C3C", "#EF5350", "#D32F2F", "#C62828", "#B71C1C", "#8B0000"
            };

            for (int i = 0; i < chips.Length; i++)
            {
                string hex = chips[i];
                Button chip = new Button();
                chip.Width = 24;
                chip.Height = 24;
                chip.Margin = new Padding(3);
                chip.FlatStyle = FlatStyle.Flat;
                chip.FlatAppearance.BorderSize = 1;
                chip.FlatAppearance.BorderColor = _theme.InputBorder;
                chip.BackColor = ParseHex(hex, Color.Gray);
                chip.Tag = hex;
                chip.Click += delegate
                {
                    if (_cmbTargetTier.SelectedItem == null)
                    {
                        return;
                    }

                    string tier = _cmbTargetTier.SelectedItem.ToString();
                    if (_colorBoxes.ContainsKey(tier))
                    {
                        _colorBoxes[tier].Text = hex;
                        UpdateSwatch(tier);
                    }
                };
                _chipPanel.Controls.Add(chip);
            }

            _colorGrid = new TableLayoutPanel();
            _colorGrid.Location = new Point(16, 226);
            _colorGrid.Size = new Size(648, 216);
            _colorGrid.ColumnCount = 4;
            _colorGrid.RowCount = 6;
            _colorGrid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 110));
            _colorGrid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 108));
            _colorGrid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 58));
            _colorGrid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 42));
            for (int r = 0; r < 6; r++)
            {
                _colorGrid.RowStyles.Add(new RowStyle(SizeType.Absolute, 34));
            }
            _colorCard.Controls.Add(_colorGrid);

            string[] tiers = new string[] { "High", "Poor", "VeryPoor", "Bad", "Offline", "Paused" };
            for (int i = 0; i < tiers.Length; i++)
            {
                string tier = tiers[i];
                Label tierLabel = new Label();
                tierLabel.Text = string.Equals(tier, "VeryPoor", StringComparison.Ordinal) ? "Very Poor" : tier;
                tierLabel.TextAlign = ContentAlignment.MiddleLeft;
                tierLabel.Dock = DockStyle.Fill;
                tierLabel.AutoEllipsis = true;
                _colorGrid.Controls.Add(tierLabel, 0, i);

                Panel swatch = new Panel();
                swatch.Width = 92;
                swatch.Height = 16;
                swatch.Margin = new Padding(6, 5, 6, 5);
                swatch.BorderStyle = BorderStyle.FixedSingle;
                swatch.Tag = "color-swatch";
                _colorGrid.Controls.Add(swatch, 1, i);
                _colorSwatches[tier] = swatch;

                TextBox txt = new TextBox();
                txt.Dock = DockStyle.Fill;
                txt.Margin = new Padding(4, 2, 8, 2);
                txt.Leave += delegate
                {
                    if (_colorBoxes.ContainsKey(tier))
                    {
                        _colorBoxes[tier].Text = NormalizeHex(_colorBoxes[tier].Text, "#808080");
                        UpdateSwatch(tier);
                    }
                };
                _colorGrid.Controls.Add(txt, 2, i);
                _colorBoxes[tier] = txt;

                Button pick = CreateSecondaryButton("Pick", 64);
                pick.Margin = new Padding(0, 0, 0, 0);
                pick.Dock = DockStyle.Left;
                pick.Click += delegate
                {
                    PickColorForTier(tier);
                };
                _colorGrid.Controls.Add(pick, 3, i);
            }

            Panel footer = CreateCardPanel();
            footer.Padding = new Padding(0);
            _rootTable.Controls.Add(footer, 0, 2);

            _footerTable = new TableLayoutPanel();
            _footerTable.Dock = DockStyle.Fill;
            _footerTable.ColumnCount = 1;
            _footerTable.RowCount = 2;
            _footerTable.Padding = new Padding(12, 8, 12, 8);
            _footerTable.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            _footerTable.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            _footerTable.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            footer.Controls.Add(_footerTable);

            _footerText = new Label();
            _footerText.AutoSize = true;
            _footerText.Text = "Changes are written to settings.json and applied immediately.";
            _footerText.ForeColor = _theme.TextSecondary;
            _footerText.Tag = "muted";
            _footerText.Margin = new Padding(2, 2, 8, 2);
            _footerTable.Controls.Add(_footerText, 0, 0);

            _footerActions = new FlowLayoutPanel();
            _footerActions.FlowDirection = FlowDirection.RightToLeft;
            _footerActions.WrapContents = true;
            _footerActions.AutoSize = false;
            _footerActions.AutoSizeMode = AutoSizeMode.GrowAndShrink;
            _footerActions.Dock = DockStyle.Fill;
            _footerActions.Margin = new Padding(0, 4, 0, 0);
            _footerTable.Controls.Add(_footerActions, 0, 1);

            Button btnProbe = CreateSecondaryButton("Probe Now", 94);
            btnProbe.Click += delegate
            {
                ForceProbe = true;
                Close();
            };

            Button btnExit = CreateSecondaryButton("Exit App", 84);
            btnExit.Click += delegate
            {
                ExitRequested = true;
                Close();
            };

            Button btnClose = CreateSecondaryButton("Close", 78);
            btnClose.Click += delegate
            {
                Close();
            };

            Button btnSave = CreatePrimaryButton("Save", 84);
            btnSave.Click += SaveAndClose;

            _footerActions.Controls.Add(btnSave);
            _footerActions.Controls.Add(btnClose);
            _footerActions.Controls.Add(btnExit);
            _footerActions.Controls.Add(btnProbe);

            WireThresholdEvents();

            Shown += delegate
            {
                FitToWorkingArea();
                ApplyResponsiveLayout();
                if (_leftColumn != null) { _leftColumn.AutoScrollPosition = new Point(0, 0); }
                if (_rightColumn != null) { _rightColumn.AutoScrollPosition = new Point(0, 0); }
            };
            Resize += delegate
            {
                ApplyResponsiveLayout();
            };
        }

        private void BindInitialValues()
        {
            _chkStartup.Checked = _initialStartup;
            _chkPaused.Checked = _initialPaused;

            _trkHigh.Value = ClampInt(_workingConfig.QualityHighMinScore, 1, 100);
            _trkPoor.Value = ClampInt(_workingConfig.QualityPoorMinScore, 0, 99);
            _trkVeryPoor.Value = ClampInt(_workingConfig.QualityVeryPoorMinScore, 0, 98);

            _numHigh.Value = _trkHigh.Value;
            _numPoor.Value = _trkPoor.Value;
            _numVeryPoor.Value = _trkVeryPoor.Value;

            _colorBoxes["High"].Text = NormalizeHex(_workingConfig.ColorHighHex, "#2ECC71");
            _colorBoxes["Poor"].Text = NormalizeHex(_workingConfig.ColorPoorHex, "#F1C40F");
            _colorBoxes["VeryPoor"].Text = NormalizeHex(_workingConfig.ColorVeryPoorHex, "#E67E22");
            _colorBoxes["Bad"].Text = NormalizeHex(_workingConfig.ColorBadHex, "#E74C3C");
            _colorBoxes["Offline"].Text = NormalizeHex(_workingConfig.ColorOfflineHex, "#951111");
            _colorBoxes["Paused"].Text = NormalizeHex(_workingConfig.ColorPausedHex, "#A0A0A0");

            UpdateSwatch("High");
            UpdateSwatch("Poor");
            UpdateSwatch("VeryPoor");
            UpdateSwatch("Bad");
            UpdateSwatch("Offline");
            UpdateSwatch("Paused");

            if (_snapshot != null)
            {
                _lblLive.Text = string.Format(
                    "{0} | Score {1:0}%\nD {2:0.0} Mbps | U {3:0.0} Mbps | L {4}",
                    _snapshot.Tier,
                    _snapshot.QualityScore,
                    _snapshot.DownloadMbps,
                    _snapshot.UploadMbps,
                    double.IsNaN(_snapshot.LatencyMs) ? "n/a" : (_snapshot.LatencyMs.ToString("0") + " ms")
                );
            }
            else
            {
                _lblLive.Text = "Waiting for first probe...";
            }

            SyncThresholdText();
        }

        private void WireThresholdEvents()
        {
            _trkHigh.ValueChanged += delegate
            {
                if (_syncingThresholds) return;
                _numHigh.Value = _trkHigh.Value;
                CoerceThresholds();
            };
            _trkPoor.ValueChanged += delegate
            {
                if (_syncingThresholds) return;
                _numPoor.Value = _trkPoor.Value;
                CoerceThresholds();
            };
            _trkVeryPoor.ValueChanged += delegate
            {
                if (_syncingThresholds) return;
                _numVeryPoor.Value = _trkVeryPoor.Value;
                CoerceThresholds();
            };

            _numHigh.ValueChanged += delegate
            {
                if (_syncingThresholds) return;
                _trkHigh.Value = (int)_numHigh.Value;
                CoerceThresholds();
            };
            _numPoor.ValueChanged += delegate
            {
                if (_syncingThresholds) return;
                _trkPoor.Value = (int)_numPoor.Value;
                CoerceThresholds();
            };
            _numVeryPoor.ValueChanged += delegate
            {
                if (_syncingThresholds) return;
                _trkVeryPoor.Value = (int)_numVeryPoor.Value;
                CoerceThresholds();
            };
        }

        private void CoerceThresholds()
        {
            _syncingThresholds = true;
            try
            {
                int high = (int)_numHigh.Value;
                int poor = (int)_numPoor.Value;
                int veryPoor = (int)_numVeryPoor.Value;

                if (poor >= high)
                {
                    poor = Math.Max(0, high - 1);
                }
                if (veryPoor >= poor)
                {
                    veryPoor = Math.Max(0, poor - 1);
                }

                high = ClampInt(high, 1, 100);
                poor = ClampInt(poor, 0, 99);
                veryPoor = ClampInt(veryPoor, 0, 98);

                _numHigh.Value = high;
                _numPoor.Value = poor;
                _numVeryPoor.Value = veryPoor;

                _trkHigh.Value = high;
                _trkPoor.Value = poor;
                _trkVeryPoor.Value = veryPoor;

                SyncThresholdText();
            }
            finally
            {
                _syncingThresholds = false;
            }
        }

        private void SyncThresholdText()
        {
            int high = (int)_numHigh.Value;
            int poor = (int)_numPoor.Value;
            int veryPoor = (int)_numVeryPoor.Value;

            _lblRanges.Text = string.Format(
                "High: {0}-100 | Poor: {1}-{2} | Very Poor: {3}-{4} | Bad: 0-{5}",
                high,
                poor,
                Math.Max(high - 1, poor),
                veryPoor,
                Math.Max(poor - 1, veryPoor),
                Math.Max(veryPoor - 1, 0)
            );
        }

        private static int ClampInt(int value, int min, int max)
        {
            if (value < min) return min;
            if (value > max) return max;
            return value;
        }

        private static int GetButtonWidth(string text, Font font, int minimumWidth)
        {
            string safe = text ?? string.Empty;
            Font effective = font ?? SystemFonts.MessageBoxFont;
            Size measured = TextRenderer.MeasureText(
                safe,
                effective,
                new Size(int.MaxValue, int.MaxValue),
                TextFormatFlags.SingleLine | TextFormatFlags.NoPrefix | TextFormatFlags.NoPadding);
            int width = measured.Width + 26;
            return Math.Max(minimumWidth, width);
        }

        private static string NormalizeHex(string value, string fallback)
        {
            string raw = string.IsNullOrWhiteSpace(value) ? fallback : value.Trim();
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

        private static Color ParseHex(string value, Color fallback)
        {
            string hex = NormalizeHex(value, "#808080");
            try
            {
                return ColorTranslator.FromHtml(hex);
            }
            catch
            {
                return fallback;
            }
        }

        private void UpdateSwatch(string tier)
        {
            if (!_colorBoxes.ContainsKey(tier) || !_colorSwatches.ContainsKey(tier))
            {
                return;
            }

            string normalized = NormalizeHex(_colorBoxes[tier].Text, "#808080");
            _colorBoxes[tier].Text = normalized;
            _colorSwatches[tier].BackColor = ParseHex(normalized, Color.Gray);
        }

        private void PickColorForTier(string tier)
        {
            if (!_colorBoxes.ContainsKey(tier))
            {
                return;
            }

            _colorDialog.FullOpen = true;
            _colorDialog.Color = ParseHex(_colorBoxes[tier].Text, Color.Gray);
            if (_colorDialog.ShowDialog() == DialogResult.OK)
            {
                string hex = string.Format("#{0:X2}{1:X2}{2:X2}", _colorDialog.Color.R, _colorDialog.Color.G, _colorDialog.Color.B);
                _colorBoxes[tier].Text = hex;
                UpdateSwatch(tier);
            }
        }

        private void SaveAndClose(object sender, EventArgs e)
        {
            try
            {
                CoerceThresholds();
                _workingConfig.QualityHighMinScore = (int)_numHigh.Value;
                _workingConfig.QualityPoorMinScore = (int)_numPoor.Value;
                _workingConfig.QualityVeryPoorMinScore = (int)_numVeryPoor.Value;

                _workingConfig.ColorHighHex = NormalizeHex(_colorBoxes["High"].Text, "#2ECC71");
                _workingConfig.ColorPoorHex = NormalizeHex(_colorBoxes["Poor"].Text, "#F1C40F");
                _workingConfig.ColorVeryPoorHex = NormalizeHex(_colorBoxes["VeryPoor"].Text, "#E67E22");
                _workingConfig.ColorBadHex = NormalizeHex(_colorBoxes["Bad"].Text, "#E74C3C");
                _workingConfig.ColorOfflineHex = NormalizeHex(_colorBoxes["Offline"].Text, "#951111");
                _workingConfig.ColorPausedHex = NormalizeHex(_colorBoxes["Paused"].Text, "#A0A0A0");
                _workingConfig.Normalize();

                UpdatedConfig = _workingConfig;
                StartupEnabled = _chkStartup.Checked;
                Paused = _chkPaused.Checked;
                ForceProbe = true;
                Saved = true;
                DialogResult = DialogResult.OK;
                Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Failed to save settings: " + ex.Message,
                    _appName,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
            }
        }

        private void ShowSnapshotDialog()
        {
            if (_snapshot == null)
            {
                MessageBox.Show(
                    "No samples collected yet.",
                    _appName,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information
                );
                return;
            }

            string latency = double.IsNaN(_snapshot.LatencyMs) ? "n/a" : _snapshot.LatencyMs.ToString("0.0") + " ms";
            string message = string.Join(Environment.NewLine, new string[]
            {
                "Interface: " + _snapshot.InterfaceName + " [" + _snapshot.InterfaceType + "]",
                "Link speed: " + _snapshot.LinkMbps.ToString("0") + " Mbps",
                "",
                "Quality: " + _snapshot.QualityScore.ToString("0") + "% (" + _snapshot.Tier + ")",
                "Download: " + _snapshot.DownloadMbps.ToString("0.00") + " Mbps",
                "Upload: " + _snapshot.UploadMbps.ToString("0.00") + " Mbps",
                "Latency: " + latency,
                "Jitter: " + _snapshot.JitterMs.ToString("0.0") + " ms",
                "Packet Loss: " + _snapshot.LossPct.ToString("0.0") + "%",
                "Consistency: " + (_snapshot.ConsistencyScore * 100.0).ToString("0.0") + "%",
                "Latency Host: " + _snapshot.LatencyHost
            });

            MessageBox.Show(message, _appName, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private Panel CreateCardPanel()
        {
            Panel panel = new Panel();
            panel.BackColor = _theme != null ? _theme.CardBackground : Color.White;
            panel.BorderStyle = BorderStyle.FixedSingle;
            panel.Tag = "card";
            panel.Margin = new Padding(0, 0, 0, 8);
            return panel;
        }

        private void AddCardHeading(Control card, string title, string subtitle)
        {
            Label titleLabel = new Label();
            titleLabel.AutoSize = true;
            titleLabel.Font = new Font("Segoe UI Variable Text", 12f, FontStyle.Bold, GraphicsUnit.Point);
            titleLabel.Text = title;
            titleLabel.ForeColor = _theme.TextPrimary;
            titleLabel.Location = new Point(14, 14);
            titleLabel.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            card.Controls.Add(titleLabel);

            Label subtitleLabel = new Label();
            subtitleLabel.AutoSize = true;
            subtitleLabel.AutoEllipsis = false;
            subtitleLabel.ForeColor = _theme.TextSecondary;
            subtitleLabel.Text = subtitle;
            subtitleLabel.Tag = "muted";
            subtitleLabel.Name = "CardSubtitleLabel";
            subtitleLabel.Location = new Point(16, 52);
            subtitleLabel.MaximumSize = new Size(Math.Max(120, card.ClientSize.Width - 30), 0);
            subtitleLabel.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            card.Controls.Add(subtitleLabel);
        }

        private Button CreateSecondaryButton(string text, int width)
        {
            Button button = new Button();
            button.Text = text;
            button.Width = GetButtonWidth(text, button.Font, width);
            button.Height = Math.Max(32, button.Font.Height + 14);
            button.FlatStyle = FlatStyle.Flat;
            button.FlatAppearance.BorderSize = 1;
            button.FlatAppearance.BorderColor = _theme.SecondaryButtonBorder;
            button.BackColor = _theme.SecondaryButtonBackground;
            button.ForeColor = _theme.SecondaryButtonText;
            button.AutoEllipsis = false;
            button.TextAlign = ContentAlignment.MiddleCenter;
            button.Padding = new Padding(6, 0, 6, 0);
            button.Tag = "secondary-button";
            button.Margin = new Padding(0, 0, 8, 0);
            return button;
        }

        private Button CreatePrimaryButton(string text, int width)
        {
            Button button = CreateSecondaryButton(text, width);
            button.BackColor = _theme.PrimaryButtonBackground;
            button.ForeColor = _theme.PrimaryButtonText;
            button.FlatAppearance.BorderColor = _theme.PrimaryButtonBorder;
            button.Tag = "primary-button";
            return button;
        }

        private static Bitmap BuildLogoBitmap(int width, int height)
        {
            Bitmap bmp = new Bitmap(width, height);
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.Clear(Color.Transparent);

                Rectangle rect = new Rectangle(2, 2, width - 4, height - 4);
                using (GraphicsPath path = RoundedRectangle(rect, 12))
                using (LinearGradientBrush brush = new LinearGradientBrush(rect, Color.FromArgb(0, 95, 184), Color.FromArgb(57, 175, 255), 80f))
                using (Pen border = new Pen(Color.FromArgb(0, 86, 170), 1.4f))
                {
                    g.FillPath(brush, path);
                    g.DrawPath(border, path);
                }

                int[] heights = new int[] { 10, 15, 22, 29 };
                for (int i = 0; i < heights.Length; i++)
                {
                    int x = 11 + (i * 9);
                    int h = heights[i];
                    int y = 42 - h;
                    using (SolidBrush b = new SolidBrush(Color.FromArgb(240, 249, 255)))
                    {
                        g.FillRectangle(b, new Rectangle(x, y, 6, h));
                    }
                }
            }
            return bmp;
        }

        private static GraphicsPath RoundedRectangle(Rectangle bounds, int radius)
        {
            int d = radius * 2;
            GraphicsPath path = new GraphicsPath();
            path.AddArc(bounds.Left, bounds.Top, d, d, 180, 90);
            path.AddArc(bounds.Right - d, bounds.Top, d, d, 270, 90);
            path.AddArc(bounds.Right - d, bounds.Bottom - d, d, d, 0, 90);
            path.AddArc(bounds.Left, bounds.Bottom - d, d, d, 90, 90);
            path.CloseFigure();
            return path;
        }

        private void AddThresholdRow(
            Control parent,
            string label,
            int y,
            out TrackBar trackBar,
            out NumericUpDown numeric,
            int min,
            int max)
        {
            Label lbl = new Label();
            lbl.Text = label;
            lbl.AutoSize = true;
            lbl.ForeColor = _theme.TextPrimary;
            lbl.Location = new Point(16, y + 7);
            parent.Controls.Add(lbl);

            trackBar = new TrackBar();
            trackBar.Minimum = min;
            trackBar.Maximum = max;
            trackBar.TickFrequency = 1;
            trackBar.AutoSize = false;
            trackBar.Width = 430;
            trackBar.Height = 40;
            trackBar.Location = new Point(136, y);
            parent.Controls.Add(trackBar);

            numeric = new NumericUpDown();
            numeric.Minimum = min;
            numeric.Maximum = max;
            numeric.Width = 82;
            numeric.Height = 34;
            numeric.Location = new Point(580, y + 4);
            numeric.TextAlign = HorizontalAlignment.Center;
            numeric.BackColor = _theme.InputBackground;
            numeric.ForeColor = _theme.TextPrimary;
            parent.Controls.Add(numeric);
        }

        private void SettingsForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            SystemEvents.UserPreferenceChanged -= SystemEvents_UserPreferenceChanged;
        }

        private void SystemEvents_UserPreferenceChanged(object sender, UserPreferenceChangedEventArgs e)
        {
            if (IsDisposed)
            {
                return;
            }

            try
            {
                BeginInvoke((Action)(() =>
                {
                    if (IsDisposed)
                    {
                        return;
                    }
                    ApplyTheme();
                    ApplyResponsiveLayout();
                }));
            }
            catch
            {
            }
        }

        private void FitToWorkingArea()
        {
            Rectangle workArea = Screen.FromControl(this).WorkingArea;

            int maxWidth = Math.Max(680, workArea.Width - 24);
            int maxHeight = Math.Max(520, workArea.Height - 24);
            int targetWidth = Math.Min(maxWidth, Math.Max(980, Math.Min(1240, maxWidth)));
            int targetHeight = Math.Min(maxHeight, Math.Max(720, Math.Min(820, maxHeight)));

            int left = workArea.Left + Math.Max(0, (workArea.Width - targetWidth) / 2);
            int top = workArea.Top + Math.Max(0, (workArea.Height - targetHeight) / 2);
            SetDesktopBounds(left, top, targetWidth, targetHeight);
        }

        private void ApplyTheme()
        {
            _theme = UiTheme.GetCurrent();
            BackColor = _theme.AppBackground;
            ForeColor = _theme.TextPrimary;
            NativeUi.ApplyWindowTheme(this, _theme.IsDark);

            ApplyThemeRecursive(this);

            if (_chipPanel != null)
            {
                _chipPanel.BackColor = _theme.ChipSurface;
            }
            if (_lblRanges != null)
            {
                _lblRanges.ForeColor = _theme.TextSecondary;
            }
            if (_footerText != null)
            {
                _footerText.ForeColor = _theme.TextSecondary;
            }
        }

        private void ApplyThemeRecursive(Control control)
        {
            if (control == null)
            {
                return;
            }

            if (control is Panel)
            {
                if (Equals(control.Tag, "card"))
                {
                    control.BackColor = _theme.CardBackground;
                    control.ForeColor = _theme.TextPrimary;
                }
                else if (Equals(control.Tag, "color-swatch"))
                {
                    control.ForeColor = _theme.TextPrimary;
                }
                else
                {
                    control.BackColor = _theme.AppBackground;
                    control.ForeColor = _theme.TextPrimary;
                }
            }
            else if (control is Label)
            {
                control.BackColor = Color.Transparent;
                control.ForeColor = Equals(control.Tag, "muted") ? _theme.TextSecondary : _theme.TextPrimary;
            }
            else if (control is CheckBox)
            {
                control.ForeColor = _theme.TextPrimary;
                control.BackColor = _theme.CardBackground;
            }
            else if (control is TextBox)
            {
                control.BackColor = _theme.InputBackground;
                control.ForeColor = _theme.TextPrimary;
            }
            else if (control is ComboBox)
            {
                control.BackColor = _theme.InputBackground;
                control.ForeColor = _theme.TextPrimary;
            }
            else if (control is NumericUpDown)
            {
                control.BackColor = _theme.InputBackground;
                control.ForeColor = _theme.TextPrimary;
            }
            else if (control is TrackBar)
            {
                control.BackColor = _theme.CardBackground;
                control.ForeColor = _theme.TextPrimary;
            }
            else if (control is Button)
            {
                Button button = (Button)control;
                if (Equals(button.Tag, "primary-button"))
                {
                    button.BackColor = _theme.PrimaryButtonBackground;
                    button.ForeColor = _theme.PrimaryButtonText;
                    button.FlatAppearance.BorderColor = _theme.PrimaryButtonBorder;
                }
                else if (!Regex.IsMatch(button.Tag as string ?? string.Empty, "^#[0-9A-Fa-f]{6}$"))
                {
                    button.BackColor = _theme.SecondaryButtonBackground;
                    button.ForeColor = _theme.SecondaryButtonText;
                    button.FlatAppearance.BorderColor = _theme.SecondaryButtonBorder;
                }
            }
            else if (control is FlowLayoutPanel || control is TableLayoutPanel || control is SplitContainer)
            {
                bool insideCard = control.Parent != null && Equals(control.Parent.Tag, "card");
                control.BackColor = insideCard ? _theme.CardBackground : _theme.AppBackground;
                control.ForeColor = _theme.TextPrimary;
            }

            foreach (Control child in control.Controls)
            {
                ApplyThemeRecursive(child);
            }
        }

        private void ApplyResponsiveLayout()
        {
            if (_headerCard != null && _headerTitle != null && _headerSubtitle != null)
            {
                int textWidth = Math.Max(160, _headerCard.ClientSize.Width - 102);
                _headerTitle.MaximumSize = new Size(textWidth, 0);
                _headerSubtitle.MaximumSize = new Size(textWidth, 0);
                _headerSubtitle.Size = new Size(textWidth, 24);
            }

            if (_bodySplit != null)
            {
                int splitWidth = Math.Max(0, _bodySplit.ClientSize.Width);
                bool compact = splitWidth < 1220;
                if (compact)
                {
                    _bodySplit.FixedPanel = FixedPanel.None;
                    _bodySplit.Orientation = Orientation.Horizontal;
                    _bodySplit.Panel1MinSize = 230;
                    _bodySplit.Panel2MinSize = 280;
                    int availableHeight = Math.Max(0, _bodySplit.ClientSize.Height - _bodySplit.SplitterWidth);
                    int desiredTop = Math.Max(250, Math.Min(380, availableHeight / 3));
                    int maxTop = Math.Max(_bodySplit.Panel1MinSize, availableHeight - _bodySplit.Panel2MinSize);
                    _bodySplit.SplitterDistance = Math.Max(_bodySplit.Panel1MinSize, Math.Min(desiredTop, maxTop));
                }
                else
                {
                    _bodySplit.Orientation = Orientation.Vertical;
                    _bodySplit.FixedPanel = FixedPanel.Panel1;
                    ApplySplitLayout(_bodySplit, 390, 300, 380);
                }
            }

            if (_leftColumn != null)
            {
                int minLeftWidth = (_bodySplit != null && _bodySplit.Orientation == Orientation.Horizontal) ? 220 : 290;
                int leftWidth = Math.Max(minLeftWidth, _leftColumn.ClientSize.Width - SystemInformation.VerticalScrollBarWidth - 8);
                foreach (Control c in _leftColumn.Controls)
                {
                    c.Width = leftWidth;
                    for (int i = 0; i < c.Controls.Count; i++)
                    {
                        Label subtitle = c.Controls[i] as Label;
                        if (subtitle != null && string.Equals(subtitle.Name, "CardSubtitleLabel", StringComparison.Ordinal))
                        {
                            subtitle.MaximumSize = new Size(Math.Max(120, c.ClientSize.Width - 30), 0);
                        }
                    }
                }
            }

            if (_rightColumn != null)
            {
                int minRightWidth = (_bodySplit != null && _bodySplit.Orientation == Orientation.Horizontal) ? 260 : 360;
                int rightWidth = Math.Max(minRightWidth, _rightColumn.ClientSize.Width - SystemInformation.VerticalScrollBarWidth - 10);
                foreach (Control c in _rightColumn.Controls)
                {
                    c.Width = rightWidth;
                    for (int i = 0; i < c.Controls.Count; i++)
                    {
                        Label subtitle = c.Controls[i] as Label;
                        if (subtitle != null && string.Equals(subtitle.Name, "CardSubtitleLabel", StringComparison.Ordinal))
                        {
                            subtitle.MaximumSize = new Size(Math.Max(120, c.ClientSize.Width - 30), 0);
                        }
                    }
                }
            }

            RelayoutLeftCards();

            if (_thresholdCard != null && _trkHigh != null && _numHigh != null)
            {
                int contentWidth = Math.Max(260, _thresholdCard.ClientSize.Width - 28);
                int sliderLeft = 136;
                int numericWidth = 82;
                int gap = 8;
                int sliderWidth = Math.Max(150, contentWidth - sliderLeft - numericWidth - gap - 8);

                _trkHigh.Width = sliderWidth;
                _numHigh.Left = _trkHigh.Right + gap;

                _trkPoor.Width = sliderWidth;
                _numPoor.Left = _trkPoor.Right + gap;

                _trkVeryPoor.Width = sliderWidth;
                _numVeryPoor.Left = _trkVeryPoor.Right + gap;

                _lblRanges.Width = Math.Max(160, _thresholdCard.ClientSize.Width - 34);
            }

            if (_colorCard != null && _chipPanel != null && _colorGrid != null)
            {
                int width = Math.Max(280, _colorCard.ClientSize.Width - 32);
                _chipPanel.Width = width;
                _colorGrid.Width = width;
                _colorGrid.ColumnStyles[0].SizeType = SizeType.Absolute;
                _colorGrid.ColumnStyles[0].Width = 140;
                _colorGrid.ColumnStyles[1].SizeType = SizeType.Absolute;
                _colorGrid.ColumnStyles[1].Width = 110;
                _colorGrid.ColumnStyles[2].SizeType = SizeType.Percent;
                _colorGrid.ColumnStyles[2].Width = 65;
                _colorGrid.ColumnStyles[3].SizeType = SizeType.Percent;
                _colorGrid.ColumnStyles[3].Width = 35;
                for (int row = 0; row < _colorGrid.RowStyles.Count; row++)
                {
                    _colorGrid.RowStyles[row].SizeType = SizeType.Absolute;
                    _colorGrid.RowStyles[row].Height = Math.Max(36, Font.Height + 16);
                }
                _colorGrid.Top = _chipPanel.Bottom + 10;
                _colorCard.Height = Math.Max(470, _colorGrid.Bottom + 14);
            }

            if (_footerTable != null && _footerActions != null && _footerText != null)
            {
                int available = Math.Max(220, _footerTable.ClientSize.Width - 24);
                _footerText.MaximumSize = new Size(available, 0);
                for (int i = 0; i < _footerActions.Controls.Count; i++)
                {
                    Button button = _footerActions.Controls[i] as Button;
                    if (button == null)
                    {
                        continue;
                    }

                    int minimum = Math.Max(76, button.Width);
                    button.Width = GetButtonWidth(button.Text, button.Font, minimum);
                }

                _footerActions.Width = available;
                Size preferred = _footerActions.GetPreferredSize(new Size(available, 0));
                _footerActions.Height = Math.Max(36, preferred.Height);

                if (_rootTable != null && _rootTable.RowStyles.Count >= 3)
                {
                    Size footerPreferred = _footerTable.GetPreferredSize(new Size(Math.Max(240, _footerTable.ClientSize.Width), 0));
                    int rowHeight = Math.Max(124, footerPreferred.Height + 20);
                    rowHeight = Math.Min(220, rowHeight);
                    _rootTable.RowStyles[2].SizeType = SizeType.Absolute;
                    _rootTable.RowStyles[2].Height = rowHeight;
                }
            }

            AdjustButtonTextClipping(this);
        }

        private void RelayoutLeftCards()
        {
            if (_leftColumn == null)
            {
                return;
            }

            for (int i = 0; i < _leftColumn.Controls.Count; i++)
            {
                Panel card = _leftColumn.Controls[i] as Panel;
                if (card == null)
                {
                    continue;
                }

                Label subtitle = null;
                for (int child = 0; child < card.Controls.Count; child++)
                {
                    Label label = card.Controls[child] as Label;
                    if (label != null && string.Equals(label.Name, "CardSubtitleLabel", StringComparison.Ordinal))
                    {
                        subtitle = label;
                        break;
                    }
                }

                int contentTop = subtitle != null ? subtitle.Bottom + 10 : 96;
                if (string.Equals(card.Name, "LiveCard", StringComparison.Ordinal) && _lblLive != null)
                {
                    _lblLive.Left = 14;
                    _lblLive.Top = contentTop;
                    _lblLive.Width = Math.Max(160, card.ClientSize.Width - 28);
                    _lblLive.Height = Math.Max(52, card.ClientSize.Height - contentTop - 14);
                    card.Height = Math.Max(176, _lblLive.Bottom + 12);
                }
                else if (string.Equals(card.Name, "GeneralCard", StringComparison.Ordinal) && _chkStartup != null && _chkPaused != null)
                {
                    _chkStartup.Left = 16;
                    _chkStartup.Top = contentTop;
                    _chkPaused.Left = 16;
                    _chkPaused.Top = _chkStartup.Bottom + 8;
                    card.Height = Math.Max(176, _chkPaused.Bottom + 14);
                }
                else if (string.Equals(card.Name, "ActionsCard", StringComparison.Ordinal) && _btnStatusDetails != null)
                {
                    _btnStatusDetails.Left = 16;
                    _btnStatusDetails.Top = contentTop;
                    card.Height = Math.Max(146, _btnStatusDetails.Bottom + 14);
                }
            }
        }

        private void AdjustButtonTextClipping(Control control)
        {
            if (control == null)
            {
                return;
            }

            Button button = control as Button;
            if (button != null)
            {
                string tag = button.Tag as string ?? string.Empty;
                bool isColorChip = Regex.IsMatch(tag, "^#[0-9A-Fa-f]{6}$");
                if (!isColorChip)
                {
                    button.AutoEllipsis = false;
                    button.TextAlign = ContentAlignment.MiddleCenter;
                    button.Height = Math.Max(button.Height, button.Font.Height + 14);
                    if (button.Dock == DockStyle.None || button.Dock == DockStyle.Left || button.Dock == DockStyle.Right)
                    {
                        int minimum = Math.Max(64, button.Width);
                        button.Width = GetButtonWidth(button.Text, button.Font, minimum);
                    }
                }
            }

            foreach (Control child in control.Controls)
            {
                AdjustButtonTextClipping(child);
            }
        }

        private static void ApplySplitLayout(SplitContainer container, int preferred, int desiredMinLeft, int desiredMinRight)
        {
            if (container == null)
            {
                return;
            }

            int width = container.ClientSize.Width;
            if (width <= 0)
            {
                return;
            }

            int splitter = Math.Max(container.SplitterWidth, 4);
            int available = width - splitter;
            if (available < 100)
            {
                return;
            }

            int minLeft = desiredMinLeft;
            int minRight = desiredMinRight;
            if (minLeft + minRight > available)
            {
                double ratio = desiredMinLeft / (double)Math.Max(1, desiredMinLeft + desiredMinRight);
                minLeft = Math.Max(60, (int)Math.Round(available * ratio));
                minRight = Math.Max(60, available - minLeft);
            }

            minLeft = Math.Max(40, Math.Min(minLeft, available - 40));
            minRight = Math.Max(40, Math.Min(minRight, available - minLeft));

            try
            {
                container.Panel1MinSize = minLeft;
                container.Panel2MinSize = minRight;
            }
            catch
            {
                try
                {
                    container.Panel1MinSize = 40;
                    container.Panel2MinSize = 40;
                }
                catch
                {
                }
            }

            int min = Math.Max(0, container.Panel1MinSize);
            int max = width - container.Panel2MinSize - splitter;
            if (max < min)
            {
                return;
            }

            int safe = Math.Max(min, Math.Min(preferred, max));
            try
            {
                if (container.SplitterDistance != safe)
                {
                    container.SplitterDistance = safe;
                }
            }
            catch
            {
            }
        }
    }
}
