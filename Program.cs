using System.Collections.Concurrent;
using System.Diagnostics;
using System.ComponentModel;
using System.Drawing.Drawing2D;
using Microsoft.Win32;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Unicode;
using System.Threading;

namespace KeyMouseHeatmap;

internal static class Program
{
    [STAThread]
    private static void Main()
    {
        ApplicationConfiguration.Initialize();
        Application.Run(new MainForm());
    }
}

internal sealed class MainForm : Form
{
    private const string AppVersion = "v1.8.1";
    private const int WmDrainInput = 0x8000 + 0x421;
    private readonly string appDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "KeyMouseHeatmap", "data");
    private readonly JsonSerializerOptions jsonOptions = new() { WriteIndented = false, Encoder = JavaScriptEncoder.Create(UnicodeRanges.All) };
    private readonly string selfProcessName = Process.GetCurrentProcess().ProcessName;
    private readonly Icon appIcon = Icon.ExtractAssociatedIcon(Application.ExecutablePath) ?? SystemIcons.Application;
    private readonly CombinedView combinedView = new();
    private readonly KeyboardView keyboardView = new();
    private readonly MouseView mouseView = new();
    private readonly PeakView peakView = new();
    private readonly SpeedView speedView = new();
    private readonly ListView rankList = new();
    private readonly ListView groupList = new();
    private readonly ListView appDetailList = new();
    private readonly ImageList appIconList = new() { ImageSize = new Size(20, 20), ColorDepth = ColorDepth.Depth32Bit };
    private readonly TabControl tabs = new HeatmapTabControl();
    private readonly StatusStrip statusStrip = new();
    private readonly ToolStripStatusLabel stateLabel = new();
    private readonly ToolStripStatusLabel totalLabel = new();
    private readonly ToolStripStatusLabel pathLabel = new() { Spring = true, TextAlign = ContentAlignment.MiddleRight };
    private readonly ToolStripButton startButton = new("开始");
    private readonly ToolStripButton todayButton = new("今天");
    private readonly ToolStripButton prevButton = new("前一天");
    private readonly ToolStripButton nextButton = new("后一天");
    private readonly ToolStripButton resetButton = new("清空");
    private readonly ToolStripButton exportCsvButton = new("CSV");
    private readonly ToolStripButton exportPngButton = new("图片");
    private readonly ToolStripButton openDataButton = new("数据目录");
    private readonly ToolStripButton topMostButton = new("置顶") { CheckOnClick = true };
    private readonly ToolStripButton trayButton = new("托盘") { CheckOnClick = true };
    private readonly ToolStripButton minimizeBackgroundButton = new("最小化后台") { CheckOnClick = true };
    private readonly ToolStripButton autoStartButton = new("自动开始") { CheckOnClick = true, Checked = true };
    private readonly ToolStripButton exportHtmlButton = new("报告");
    private readonly ToolStripComboBox trackModeBox = new() { DropDownStyle = ComboBoxStyle.DropDownList, Width = 108 };
    private readonly ToolStripComboBox rangeBox = new() { DropDownStyle = ComboBoxStyle.DropDownList, Width = 118 };
    private readonly ToolStripComboBox themeBox = new() { DropDownStyle = ComboBoxStyle.DropDownList, Width = 96 };
    private readonly ToolStripButton startupButton = new("开机启动") { CheckOnClick = true };
    private readonly ToolStripButton overlayButton = new("悬浮窗") { CheckOnClick = true };
    private readonly DateTimePicker datePicker = new() { Format = DateTimePickerFormat.Custom, CustomFormat = "yyyy-MM-dd", Width = 112 };
    private readonly OverlayForm overlay = new();
    private readonly HashSet<string> physicalDownKeys = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, DateTime> physicalDownLastSeen = new(StringComparer.OrdinalIgnoreCase);
    private readonly HashSet<string> physicalDownMouse = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, DateTime> mouseMomentaryUntil = new(StringComparer.OrdinalIgnoreCase);
    private readonly ConcurrentQueue<InputEvent> pendingInputs = new();
    private readonly Queue<DateTime> recentKeyTimestamps = new();
    private readonly object saveLock = new();
    private static readonly TimeSpan PhysicalKeyStaleTimeout = TimeSpan.FromSeconds(8);
    private static readonly TimeSpan MouseMomentaryHighlight = TimeSpan.FromMilliseconds(80);
    private static readonly TimeSpan AppUsageUiRefreshInterval = TimeSpan.FromSeconds(1);
    private readonly NotifyIcon trayIcon = new();
    private readonly System.Windows.Forms.Timer refreshTimer = new();
    private readonly System.Windows.Forms.Timer liveRefreshTimer = new();
    private readonly System.Windows.Forms.Timer mouseDistanceTimer = new();
    private readonly System.Windows.Forms.Timer appUsageTimer = new();
    private readonly System.Windows.Forms.Timer resizeTimer = new();
    private readonly System.Windows.Forms.Timer saveTimer = new();
    private readonly System.Windows.Forms.Timer settingsSaveTimer = new();
    private UsageState state;
    private UsageState displayState = new();
    private AppSettings settings = new();
    private DateOnly selectedDate = DateOnly.FromDateTime(DateTime.Today);
    private bool isTracking;
    private bool liveDirty;
    private bool pressedDirty;
    private bool isResizing;
    private bool isInteractiveMoveOrResize;
    private int keyTotalCache;
    private int mouseTotalCache;
    private bool rankDirty = true;
    private bool groupDirty = true;
    private bool appDetailDirty = true;
    private bool stateDirty;
    private bool settingsDirty;
    private volatile bool saveInProgress;
    private int inputDrainScheduled;
    private long lastImmediateVisualUpdateTick;
    private bool overlayKeyDirty;
    private string overlayKeyName = string.Empty;
    private int overlayKeyCount;
    private readonly ConcurrentDictionary<uint, byte> pendingAppPathResolves = new();
    private Point? lastMousePoint;
    private long lastDistanceUiTick;
    private DateTime? lastAppUsageTickUtc;
    private DateTime lastAppUsageUiRefreshUtc = DateTime.MinValue;
    private string lastAppUsageUiSignature = string.Empty;
    private int rankTotalCount = 1;
    private int groupTotalCount = 1;
    private int appTotalCount = 1;
    private int rankSortColumn = 3;
    private SortOrder rankSortOrder = SortOrder.Descending;
    private int groupSortColumn = 1;
    private SortOrder groupSortOrder = SortOrder.Descending;
    private int appSortColumn = 1;
    private SortOrder appSortOrder = SortOrder.Descending;
    private float listZoom = 1.0f;
    private readonly Font baseListFont = new("Microsoft YaHei UI", 9F);
    private List<RankRow> rankRows = new();
    private List<(string Group, int Count, string Note)> groupRows = new();
    private List<AppKeyRow> appKeyRows = new();

    private bool ViewingToday => selectedDate == DateOnly.FromDateTime(DateTime.Today);
    private string CurrentDataFile => DataFile(selectedDate);

    public MainForm()
    {
        Directory.CreateDirectory(appDir);
        MigrateVersionData();
        settings = LoadSettings();
        MigrateOldData();
        state = LoadState(selectedDate);
        RecalculateTotals();

        Text = "KeyMouse Heatmap";
        Icon = appIcon;
        Size = new Size(1060, 720);
        MinimumSize = new Size(980, 560);
        StartPosition = FormStartPosition.CenterScreen;
        ApplySavedWindowSettings();
        BackColor = Color.FromArgb(247, 248, 250);
        Font = new Font("Microsoft YaHei UI", 9F);

        Controls.Add(BuildTabs());
        Controls.Add(BuildToolStrip());
        Controls.Add(BuildStatusStrip());
        ConfigureTray();
        settingsSaveTimer.Interval = 750;
        settingsSaveTimer.Tick += (_, _) =>
        {
            settingsSaveTimer.Stop();
            if (settingsDirty) SaveSettings();
        };
        WireEvents();

        combinedView.StateProvider = () => displayState;
        combinedView.PressedKeysProvider = GetLivePressedKeys;
        combinedView.PressedMouseProvider = GetLivePressedMouse;
        combinedView.AppSortColumn = appSortColumn;
        combinedView.AppSortOrder = appSortOrder;
        keyboardView.StateProvider = () => displayState;
        keyboardView.PressedKeysProvider = GetLivePressedKeys;
        mouseView.StateProvider = () => displayState;
        mouseView.PressedMouseProvider = GetLivePressedMouse;
        peakView.StateProvider = () => displayState;
        speedView.StateProvider = () => displayState;
        speedView.SpeedProvider = () => BuildSpeedStats(state);
        HeatmapRenderer.LiveSpeedProvider = () => BuildSpeedStats(state);
        datePicker.Value = selectedDate.ToDateTime(TimeOnly.MinValue);
        startupButton.Checked = IsStartupEnabled();
        minimizeBackgroundButton.Checked = settings.MinimizeToBackground;
        autoStartButton.Checked = settings.AutoStart;
        trayButton.Checked = settings.TrayIconVisible;
        overlayButton.Checked = settings.OverlayVisible;
        listZoom = Math.Clamp(settings.ListZoom <= 0 ? 1.0f : settings.ListZoom, 0.82f, 1.55f);
        rankSortColumn = Math.Clamp(settings.RankSortColumn, 0, 4);
        rankSortOrder = NormalizeSortOrder(settings.RankSortOrder, SortOrder.Descending);
        groupSortColumn = Math.Clamp(settings.GroupSortColumn, 0, 3);
        groupSortOrder = NormalizeSortOrder(settings.GroupSortOrder, SortOrder.Descending);
        appSortColumn = Math.Clamp(settings.AppSortColumn, 0, 5);
        appSortOrder = NormalizeSortOrder(settings.AppSortOrder, SortOrder.Descending);
        combinedView.AppSortColumn = appSortColumn;
        combinedView.AppSortOrder = appSortOrder;
        combinedView.OverviewTopRatio = settings.OverviewTopRatio;
        combinedView.OverviewMouseRatio = settings.OverviewMouseRatio;
        combinedView.LayoutSettingsChanged += (_, _) => CaptureAndQueueSettingsSave();
        trackModeBox.Items.AddRange(new object[] { "全部统计", "只统计键盘", "只统计鼠标" });
        rangeBox.Items.AddRange(new object[] { "今天", "最近7天", "最近30天", "全部" });
        themeBox.Items.AddRange(new object[] { "浅色", "深色", "赛博蓝", "极简" });
        trackModeBox.SelectedItem = string.IsNullOrWhiteSpace(settings.TrackMode) ? "全部统计" : settings.TrackMode;
        rangeBox.SelectedItem = string.IsNullOrWhiteSpace(settings.Range) ? "今天" : settings.Range;
        themeBox.SelectedItem = string.IsNullOrWhiteSpace(settings.Theme) ? "浅色" : settings.Theme;
        ApplyTheme(themeBox.SelectedItem?.ToString() ?? "浅色");
        ApplyListZoom();
        if (settings.SelectedTabIndex >= 0 && settings.SelectedTabIndex < tabs.TabPages.Count)
            tabs.SelectedIndex = settings.SelectedTabIndex;

        refreshTimer.Interval = 100;
        refreshTimer.Tick += (_, _) =>
        {
            RefreshRankList();
            RefreshGroupStats();
            RefreshAppDetailList();
        };
        refreshTimer.Start();

        liveRefreshTimer.Interval = 10;
        liveRefreshTimer.Tick += (_, _) =>
        {
            DrainPendingInputEvents();
            // 按键高亮和统计刷新分离：
            // 1) 多个 KeyUp/KeyDown 在同一帧内合并，只重绘一次键盘；
            // 2) 松开很多键时不会一个键一个键慢慢恢复；
            // 3) 长按重复 KeyDown 不触发 UI 重绘。
            if (CleanupStalePhysicalDownKeys()) pressedDirty = true;

            if (IsHeavyUiSuspended()) return;

            FlushLiveUiAfterInput();
        };
        liveRefreshTimer.Start();

        mouseDistanceTimer.Interval = 100;
        mouseDistanceTimer.Tick += (_, _) => PollMouseDistance();

        appUsageTimer.Interval = 1000;
        appUsageTimer.Tick += (_, _) => { if (!IsHeavyUiSuspended()) PollAppUsageAsync(); };

        resizeTimer.Interval = 140;
        resizeTimer.Tick += (_, _) =>
        {
            resizeTimer.Stop();
            if (isInteractiveMoveOrResize) return;
            SetMainWindowMoveResizeState(false);
            RefreshAll();
        };

        saveTimer.Interval = 10_000;
        saveTimer.Tick += (_, _) => { if (!IsHeavyUiSuspended()) SaveState(); };
        saveTimer.Start();

        if (autoStartButton.Checked) StartTracking();
        RefreshAll();
    }

    private Control BuildToolStrip()
    {
        ApplyToolbarIcons();

        var panel = new TableLayoutPanel
        {
            Dock = DockStyle.Top,
            AutoSize = true,
            ColumnCount = 1,
            RowCount = 2,
            BackColor = Color.White,
            Padding = new Padding(0),
            Margin = new Padding(0)
        };
        panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));

        ToolStrip MakeStrip() => new()
        {
            Dock = DockStyle.Top,
            GripStyle = ToolStripGripStyle.Hidden,
            Padding = new Padding(8, 4, 8, 4),
            ImageScalingSize = new Size(18, 18),
            BackColor = Color.White,
            RenderMode = ToolStripRenderMode.ManagerRenderMode,
            Renderer = new StrongToolStripRenderer(),
            CanOverflow = false,
            AutoSize = true,
            Stretch = true
        };

        var line1 = MakeStrip();
        var line2 = MakeStrip();

        foreach (var item in new ToolStripItem[]
        {
            startButton, new ToolStripSeparator(),
            prevButton, todayButton, nextButton, new ToolStripControlHost(datePicker) { Margin = new Padding(4, 2, 6, 2) }, new ToolStripLabel("范围"), rangeBox
        })
        {
            item.Margin = new Padding(2);
            line1.Items.Add(item);
        }

        foreach (var item in new ToolStripItem[]
        {
            resetButton, exportCsvButton, exportPngButton, exportHtmlButton, openDataButton, new ToolStripSeparator(),
            topMostButton, trayButton, minimizeBackgroundButton, startupButton, overlayButton, autoStartButton, new ToolStripSeparator(), new ToolStripLabel("统计"), trackModeBox, new ToolStripLabel("主题"), themeBox
        })
        {
            item.Margin = new Padding(2);
            line2.Items.Add(item);
        }

        panel.Controls.Add(line1, 0, 0);
        panel.Controls.Add(line2, 0, 1);
        return panel;
    }

    private void ApplyToolbarIcons()
    {
        var items = new (ToolStripButton Button, string Glyph, Color Color)[]
        {
            (startButton, "▶", Color.FromArgb(28, 132, 83)),
            (prevButton, "‹", Color.FromArgb(80, 98, 120)),
            (todayButton, "●", Color.FromArgb(42, 108, 176)),
            (nextButton, "›", Color.FromArgb(80, 98, 120)),
            (resetButton, "×", Color.FromArgb(199, 74, 61)),
            (exportCsvButton, "≡", Color.FromArgb(57, 126, 172)),
            (exportPngButton, "▧", Color.FromArgb(91, 91, 176)),
            (exportHtmlButton, "◆", Color.FromArgb(42, 125, 155)),
            (openDataButton, "□", Color.FromArgb(99, 112, 128)),
            (topMostButton, "▲", Color.FromArgb(145, 96, 36)),
            (trayButton, "▁", Color.FromArgb(88, 116, 144)),
            (minimizeBackgroundButton, "↘", Color.FromArgb(88, 116, 144)),
            (startupButton, "⚙", Color.FromArgb(80, 120, 82)),
            (overlayButton, "◈", Color.FromArgb(138, 85, 160)),
            (autoStartButton, "✓", Color.FromArgb(28, 132, 83))
        };

        foreach (var item in items)
        {
            item.Button.Image = IconBitmap(item.Glyph, item.Color);
            item.Button.DisplayStyle = ToolStripItemDisplayStyle.ImageAndText;
        }
    }

    private static Bitmap IconBitmap(string text, Color color)
    {
        var bmp = new Bitmap(18, 18);
        using var g = Graphics.FromImage(bmp);
        g.Clear(Color.Transparent);
        using var font = new Font("Segoe UI Symbol", 11, FontStyle.Bold);
        TextRenderer.DrawText(g, text, font, new Rectangle(0, 0, 18, 18), color, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
        return bmp;
    }

    private StatusStrip BuildStatusStrip()
    {
        statusStrip.Dock = DockStyle.Bottom;
        statusStrip.SizingGrip = true;
        statusStrip.Items.AddRange([stateLabel, totalLabel, pathLabel]);
        return statusStrip;
    }

    private TabControl BuildTabs()
    {
        tabs.Dock = DockStyle.Fill;
        // 不让 TabControl 抢占方向键，避免按 ←/→ 时从总览跳到键盘/鼠标等页面。
        tabs.TabStop = false;
        tabs.Padding = new Point(14, 6);
        tabs.ImageList = new ImageList { ImageSize = new Size(16, 16), ColorDepth = ColorDepth.Depth32Bit };
        tabs.ImageList.Images.Add("overview", IconBitmap("◫", Color.FromArgb(48, 75, 110)));
        tabs.ImageList.Images.Add("keyboard", IconBitmap("⌨", Color.FromArgb(48, 75, 110)));
        tabs.ImageList.Images.Add("mouse", IconBitmap("◉", Color.FromArgb(48, 75, 110)));
        tabs.ImageList.Images.Add("peak", IconBitmap("▥", Color.FromArgb(48, 75, 110)));
        tabs.ImageList.Images.Add("rank", IconBitmap("≡", Color.FromArgb(48, 75, 110)));
        tabs.ImageList.Images.Add("speed", IconBitmap("↯", Color.FromArgb(48, 75, 110)));
        tabs.ImageList.Images.Add("group", IconBitmap("▦", Color.FromArgb(48, 75, 110)));

        var overviewTab = new TabPage("总览") { BackColor = Color.FromArgb(247, 248, 250), Padding = new Padding(10), ImageKey = "overview" };
        combinedView.Dock = DockStyle.Fill;
        overviewTab.Controls.Add(combinedView);

        var keyboardTab = new TabPage("键盘") { BackColor = Color.FromArgb(247, 248, 250), Padding = new Padding(10), ImageKey = "keyboard" };
        keyboardView.Dock = DockStyle.Fill;
        keyboardTab.Controls.Add(keyboardView);

        var mouseTab = new TabPage("鼠标") { BackColor = Color.FromArgb(247, 248, 250), Padding = new Padding(10), ImageKey = "mouse" };
        mouseView.Dock = DockStyle.Fill;
        mouseTab.Controls.Add(mouseView);

        var peakTab = new TabPage("峰值") { BackColor = Color.FromArgb(247, 248, 250), Padding = new Padding(10), ImageKey = "peak" };
        peakView.Dock = DockStyle.Fill;
        peakTab.Controls.Add(peakView);

        var speedTab = new TabPage("速度") { BackColor = Color.FromArgb(247, 248, 250), Padding = new Padding(10), ImageKey = "speed" };
        speedView.Dock = DockStyle.Fill;
        speedTab.Controls.Add(speedView);

        var rankTab = new TabPage("排行") { BackColor = Color.FromArgb(247, 248, 250), Padding = new Padding(10), ImageKey = "rank" };
        rankList.Dock = DockStyle.Fill;
        rankList.View = View.Details;
        rankList.FullRowSelect = true;
        rankList.GridLines = true;
        rankList.VirtualMode = true;
        rankList.OwnerDraw = true;
        rankList.HideSelection = false;
        EnableDoubleBuffer(rankList);
        rankList.Font = baseListFont;
        rankList.MouseWheel += ListViewZoomMouseWheel;
        rankList.Resize += (_, _) => FitRankColumns();
        rankList.ColumnClick += (_, e) =>
        {
            ToggleRankSort(e.Column);
            rankDirty = true;
            RefreshRankList();
        };
        rankList.Columns.Add("排行", 70, HorizontalAlignment.Right);
        rankList.Columns.Add("设备", 90);
        rankList.Columns.Add("按键", 170);
        rankList.Columns.Add("次数", 100, HorizontalAlignment.Right);
        rankList.Columns.Add("\u7edf\u8ba1\u6761", 260);
        UpdateRankColumnHeaders();
        rankList.RetrieveVirtualItem += (_, e) =>
        {
            if (e.ItemIndex < 0 || e.ItemIndex >= rankRows.Count)
            {
                e.Item = new ListViewItem("");
                return;
            }
            var row = rankRows[e.ItemIndex];
            var item = new ListViewItem(row.Rank.ToString());
            item.SubItems.Add(row.Device);
            item.SubItems.Add(row.Name);
            item.SubItems.Add(row.Count.ToString());
            item.SubItems.Add("");
            e.Item = item;
        };
        rankList.DrawColumnHeader += (_, e) => e.DrawDefault = true;
        rankList.DrawSubItem += (_, e) =>
        {
            if (e.ItemIndex < 0 || e.ItemIndex >= rankRows.Count) { e.DrawDefault = true; return; }
            if (e.ColumnIndex == 4) DrawStatBar(e.Graphics, e.Bounds, rankRows[e.ItemIndex].Count, rankTotalCount);
            else e.DrawDefault = true;
        };
        rankTab.Controls.Add(rankList);

        var groupTab = new TabPage("分组") { BackColor = Color.FromArgb(247, 248, 250), Padding = new Padding(10), ImageKey = "group" };
        groupList.Dock = DockStyle.Fill;
        groupList.View = View.Details;
        groupList.FullRowSelect = true;
        groupList.GridLines = true;
        groupList.VirtualMode = true;
        groupList.OwnerDraw = true;
        groupList.HideSelection = false;
        EnableDoubleBuffer(groupList);
        groupList.Font = baseListFont;
        groupList.MouseWheel += ListViewZoomMouseWheel;
        groupList.Resize += (_, _) => FitGroupColumns();
        groupList.ColumnClick += (_, e) =>
        {
            ToggleGroupSort(e.Column);
            groupDirty = true;
            RefreshGroupStats();
        };
        groupList.Columns.Add("分组", 160);
        groupList.Columns.Add("次数", 120, HorizontalAlignment.Right);
        groupList.Columns.Add("\u7edf\u8ba1\u6761", 260);
        groupList.Columns.Add("说明", 520);
        UpdateGroupColumnHeaders();
        groupList.RetrieveVirtualItem += (_, e) =>
        {
            if (e.ItemIndex < 0 || e.ItemIndex >= groupRows.Count)
            {
                e.Item = new ListViewItem("");
                return;
            }
            var row = groupRows[e.ItemIndex];
            var item = new ListViewItem(row.Group);
            item.SubItems.Add(row.Count.ToString());
            item.SubItems.Add("");
            item.SubItems.Add(row.Note);
            e.Item = item;
        };
        groupList.DrawColumnHeader += (_, e) => e.DrawDefault = true;
        groupList.DrawSubItem += (_, e) =>
        {
            if (e.ItemIndex < 0 || e.ItemIndex >= groupRows.Count) { e.DrawDefault = true; return; }
            if (e.ColumnIndex == 2) DrawStatBar(e.Graphics, e.Bounds, groupRows[e.ItemIndex].Count, groupTotalCount);
            else e.DrawDefault = true;
        };
        groupTab.Controls.Add(groupList);

        tabs.ImageList.Images.Add("app", IconBitmap("A", Color.FromArgb(48, 75, 110)));
        var appTab = new TabPage("\u5e94\u7528\u8be6\u60c5") { BackColor = Color.FromArgb(247, 248, 250), Padding = new Padding(10), ImageKey = "app", Name = "appDetailTab" };
        appDetailList.Dock = DockStyle.Fill;
        appDetailList.View = View.Details;
        appDetailList.FullRowSelect = true;
        appDetailList.GridLines = true;
        appDetailList.VirtualMode = true;
        appDetailList.OwnerDraw = true;
        EnableDoubleBuffer(appDetailList);
        appDetailList.HideSelection = false;
        appDetailList.Font = baseListFont;
        appDetailList.MouseWheel += ListViewZoomMouseWheel;
        appDetailList.Resize += (_, _) => FitAppColumns();
        appDetailList.ColumnClick += (_, e) =>
        {
            ToggleAppSort(e.Column);
            combinedView.AppSortColumn = appSortColumn;
            combinedView.AppSortOrder = appSortOrder;
            appDetailDirty = true;
            RefreshAppDetailList();
        };
        appIconList.Images.Add("__default", SystemIcons.Application.ToBitmap());
        appDetailList.SmallImageList = appIconList;
        appDetailList.Columns.Add("\u5e94\u7528", 220);
        appDetailList.Columns.Add("使用时长", 110, HorizontalAlignment.Right);
        appDetailList.Columns.Add("\u603b\u6b21\u6570", 100, HorizontalAlignment.Right);
        appDetailList.Columns.Add("\u6309\u952e\u79cd\u7c7b", 100, HorizontalAlignment.Right);
        appDetailList.Columns.Add("\u7edf\u8ba1\u6761", 260);
        appDetailList.Columns.Add("\u5e38\u7528\u6309\u952e", 1500);
        UpdateAppColumnHeaders();
        appDetailList.RetrieveVirtualItem += (_, e) =>
        {
            if (e.ItemIndex < 0 || e.ItemIndex >= appKeyRows.Count)
            {
                e.Item = new ListViewItem("");
                return;
            }
            var row = appKeyRows[e.ItemIndex];
            var item = new ListViewItem(row.App) { ImageKey = row.IconKey };
            item.SubItems.Add(FormatDurationForUi(row.UsageSeconds));
            item.SubItems.Add(row.Count.ToString());
            item.SubItems.Add(row.KeyTypes.ToString());
            item.SubItems.Add("");
            item.SubItems.Add(row.Summary);
            e.Item = item;
        };
        appDetailList.DrawColumnHeader += (_, e) => e.DrawDefault = true;
        appDetailList.DrawSubItem += (_, e) =>
        {
            if (e.ItemIndex < 0 || e.ItemIndex >= appKeyRows.Count) { e.DrawDefault = true; return; }
            if (e.ColumnIndex == 0) DrawAppNameCell(e.Graphics, e.Bounds, appKeyRows[e.ItemIndex], e.Item?.Selected == true);
            else if (e.ColumnIndex == 4) DrawStatBar(e.Graphics, e.Bounds, appKeyRows[e.ItemIndex].Count, appTotalCount);
            else e.DrawDefault = true;
        };
        appTab.Controls.Add(appDetailList);

        tabs.TabPages.AddRange([overviewTab, keyboardTab, mouseTab, peakTab, speedTab, rankTab, groupTab, appTab]);
        return tabs;
    }

    private void WireEvents()
    {
        startButton.Click += (_, _) => ToggleTracking();
        prevButton.Click += (_, _) => ChangeDate(selectedDate.AddDays(-1));
        nextButton.Click += (_, _) => ChangeDate(selectedDate.AddDays(1));
        todayButton.Click += (_, _) => ChangeDate(DateOnly.FromDateTime(DateTime.Today));
        resetButton.Click += (_, _) => ResetCurrentDay();
        exportCsvButton.Click += (_, _) => ExportCsv();
        exportPngButton.Click += (_, _) => ExportPng();
        exportHtmlButton.Click += (_, _) => ExportHtmlReport();
        openDataButton.Click += (_, _) => Process.Start(new ProcessStartInfo("explorer.exe", appDir) { UseShellExecute = true });
        topMostButton.CheckedChanged += (_, _) => { TopMost = topMostButton.Checked; CaptureAndQueueSettingsSave(); };
        trayButton.CheckedChanged += (_, _) => { trayIcon.Visible = trayButton.Checked || (minimizeBackgroundButton.Checked && WindowState == FormWindowState.Minimized); CaptureAndQueueSettingsSave(); };
        minimizeBackgroundButton.CheckedChanged += (_, _) => { if (!minimizeBackgroundButton.Checked && !trayButton.Checked) trayIcon.Visible = false; CaptureAndQueueSettingsSave(); };
        autoStartButton.CheckedChanged += (_, _) => CaptureAndQueueSettingsSave();
        startupButton.CheckedChanged += (_, _) => SetStartupEnabled(startupButton.Checked);
        overlayButton.CheckedChanged += (_, _) =>
        {
            if (overlayButton.Checked) overlay.Show();
            else overlay.Hide();
            CaptureAndQueueSettingsSave();
        };
        datePicker.ValueChanged += (_, _) =>
        {
            var picked = DateOnly.FromDateTime(datePicker.Value);
            if (picked != selectedDate) ChangeDate(picked);
        };
        rangeBox.SelectedIndexChanged += (_, _) => { CaptureAndQueueSettingsSave(); RefreshAll(); };
        trackModeBox.SelectedIndexChanged += (_, _) => CaptureAndQueueSettingsSave();
        themeBox.SelectedIndexChanged += (_, _) => { ApplyTheme(themeBox.SelectedItem?.ToString() ?? "浅色"); CaptureAndQueueSettingsSave(); };
        tabs.SelectedIndexChanged += (_, _) =>
        {
            if (tabs.SelectedTab?.Text == "排行")
            {
                rankDirty = true;
                RefreshRankList();
            }
            if (tabs.SelectedTab?.Name == "appDetailTab")
            {
                appDetailDirty = true;
                RefreshAppDetailList();
            }
            if (tabs.SelectedTab?.Text == "分组")
            {
                groupDirty = true;
                RefreshGroupStats();
            }
            CaptureAndQueueSettingsSave();
        };

        ResizeBegin += (_, _) => SetMainWindowMoveResizeState(true);
        ResizeEnd += (_, _) =>
        {
            resizeTimer.Stop();
            SetMainWindowMoveResizeState(false);
            CaptureAndQueueSettingsSave();
            RefreshAll();
        };

        Resize += (_, _) =>
        {
            if (isResizing)
            {
                resizeTimer.Stop();
                resizeTimer.Start();
                return;
            }
            if (minimizeBackgroundButton.Checked && WindowState == FormWindowState.Minimized)
            {
                Hide();
                trayIcon.Visible = true;
            }
        };

        Move += (_, _) => CaptureAndQueueSettingsSave();

        FormClosing += (_, _) =>
        {
            PollAppUsage();
            SaveState(force: true);
            CaptureWindowSettings();
            settingsSaveTimer.Stop();
            SaveSettings();
            RawKeyboardInput.Unregister();
            RawMouseInput.Unregister();
            overlay.Close();
            trayIcon.Visible = false;
        };

        RawKeyboardInput.KeyPressed += HandleRawKeyPressed;
        RawKeyboardInput.KeyReleased += HandleRawKeyReleased;
        RawMouseInput.MousePressed += HandleRawMousePressed;
        RawMouseInput.MouseReleased += HandleRawMouseReleased;
    }

    private void QueueInput(InputEventKind kind, string name)
    {
        if (IsDisposed || string.IsNullOrWhiteSpace(name)) return;
        pendingInputs.Enqueue(new InputEvent(kind, name));
        ScheduleInputDrain();
    }

    private void ScheduleInputDrain()
    {
        if (!IsHandleCreated || IsDisposed) return;
        if (Interlocked.Exchange(ref inputDrainScheduled, 1) == 0)
            PostMessage(Handle, WmDrainInput, IntPtr.Zero, IntPtr.Zero);
    }

    private void HandleRawKeyPressed(string name)
    {
        if (IsMouseOnlyMode() || IsDisposed || string.IsNullOrWhiteSpace(name)) return;

        physicalDownLastSeen[name] = DateTime.UtcNow;
        if (physicalDownKeys.Add(name))
        {
            QueueInput(InputEventKind.KeyDown, name);
            RefreshPressedVisualsNow();
        }
    }

    private void HandleRawKeyReleased(string name)
    {
        if (IsDisposed || string.IsNullOrWhiteSpace(name)) return;
        if (physicalDownKeys.Remove(name) | physicalDownLastSeen.Remove(name))
            RefreshPressedVisualsNow();
    }

    private void HandleRawMousePressed(string button)
    {
        if (IsKeyboardOnlyMode() || IsDisposed || string.IsNullOrWhiteSpace(button)) return;

        var isWheel = button.StartsWith("Wheel", StringComparison.OrdinalIgnoreCase);
        if (isWheel)
            mouseMomentaryUntil[button] = DateTime.UtcNow + MouseMomentaryHighlight;
        else
            mouseMomentaryUntil.Remove(button);

        if (physicalDownMouse.Add(button) || isWheel)
            RefreshPressedVisualsNow();

        QueueInput(InputEventKind.MouseDown, button);
    }

    private void HandleRawMouseReleased(string button)
    {
        if (IsDisposed || string.IsNullOrWhiteSpace(button)) return;
        if (physicalDownMouse.Remove(button) | mouseMomentaryUntil.Remove(button))
            RefreshPressedVisualsNow();
    }

    private void RefreshPressedVisualsNow()
    {
        pressedDirty = true;
        ScheduleInputDrain();
    }

    private void FlushLiveUiAfterInput()
    {
        if (IsHeavyUiSuspended()) return;

        var uiActive = IsUiActive();
        var immediateVisual = (pressedDirty || overlayKeyDirty) && ShouldRunImmediateVisualUpdate();
        if (pressedDirty)
        {
            pressedDirty = false;
            if (overlayButton.Checked)
                overlay.SetPressedKeys(GetLivePressedKeys().Concat(GetLivePressedMouse()).ToArray(), immediate: immediateVisual);
            if (uiActive)
                InvalidateActiveView(immediate: immediateVisual);
        }

        if (overlayKeyDirty && overlayButton.Checked)
        {
            overlayKeyDirty = false;
            if (!overlay.Visible) overlay.Show();
            overlay.UpdateKey(overlayKeyName, overlayKeyCount, GetLivePressedKeys().Concat(GetLivePressedMouse()).ToArray(), BuildSpeedStats(state), immediate: immediateVisual);
        }
        else
        {
            overlayKeyDirty = false;
        }

        if (!liveDirty) return;
        liveDirty = false;
        if (uiActive || overlayButton.Checked)
            RefreshLive();
    }

    private bool ShouldRunImmediateVisualUpdate()
    {
        var nowTick = Environment.TickCount64;
        if (nowTick - lastImmediateVisualUpdateTick < 16) return false;
        lastImmediateVisualUpdateTick = nowTick;
        return true;
    }

    private void DrainPendingInputEvents()
    {
        Interlocked.Exchange(ref inputDrainScheduled, 0);
        const int maxPerTick = 2000;
        var processed = 0;
        while (processed < maxPerTick && pendingInputs.TryDequeue(out var input))
        {
            processed++;
            switch (input.Kind)
            {
                case InputEventKind.KeyDown:
                    CountKeyInput(input.Name);
                    break;
                case InputEventKind.KeyUp:
                    break;
                case InputEventKind.MouseDown:
                    if (!IsKeyboardOnlyMode()) CountMouseInput(input.Name);
                    break;
                case InputEventKind.MouseUp:
                    break;
            }
        }

        if (!pendingInputs.IsEmpty)
        {
            liveDirty = true;
            ScheduleInputDrain();
        }
    }

    protected override void WndProc(ref Message m)
    {
        const int WM_ENTERSIZEMOVE = 0x0231;
        const int WM_EXITSIZEMOVE = 0x0232;

        if (m.Msg == WmDrainInput)
        {
            DrainPendingInputEvents();
            FlushLiveUiAfterInput();
            return;
        }

        if (m.Msg == WM_ENTERSIZEMOVE)
        {
            SetMainWindowMoveResizeState(true);
            base.WndProc(ref m);
            return;
        }

        if (m.Msg == WM_EXITSIZEMOVE)
        {
            base.WndProc(ref m);
            SetMainWindowMoveResizeState(false);
            RefreshAll();
            return;
        }

        if (isTracking && (RawMouseInput.ProcessMessage(m) || RawKeyboardInput.ProcessMessage(m)))
            return;
        base.WndProc(ref m);
    }

    private void SetMainWindowMoveResizeState(bool active)
    {
        isInteractiveMoveOrResize = active;
        isResizing = active;
        combinedView.IsResizing = active;
        keyboardView.IsResizing = active;
        mouseView.IsResizing = active;
        peakView.IsResizing = active;
        speedView.IsResizing = active;
    }

    private bool IsHeavyUiSuspended() => isResizing || isInteractiveMoveOrResize || overlay.IsInteractiveMoveOrResize;

    private void CountMouseInput(string button)
    {
        CountInput(() =>
        {
            EnsureTodayForInput();
            EnsureHourlyArrays();
            AddCount(state.Mouse, button);
            mouseTotalCache++;
            state.HourlyMouse[DateTime.Now.Hour]++;
            UpdateOverlay(button, state.Mouse.TryGetValue(button, out var count) ? count : 0);
        });
    }

    private bool IsKeyboardOnlyMode() => trackModeBox.SelectedIndex == 1;
    private bool IsMouseOnlyMode() => trackModeBox.SelectedIndex == 2;

    private void PollMouseDistance()
    {
        if (!isTracking || IsKeyboardOnlyMode())
        {
            lastMousePoint = null;
            return;
        }

        var point = Cursor.Position;
        if (lastMousePoint is { } last)
        {
            var dx = point.X - last.X;
            var dy = point.Y - last.Y;
            var distance = Math.Sqrt(dx * dx + dy * dy);
            if (distance > 0 && distance < 5000)
            {
                EnsureTodayForInput();
                state.MouseDistancePixels += distance;
                stateDirty = true;
                var nowTick = Environment.TickCount64;
                if ((IsUiActive() || overlayButton.Checked) && nowTick - lastDistanceUiTick >= 250)
                {
                    lastDistanceUiTick = nowTick;
                    liveDirty = true;
                }
            }
        }
        lastMousePoint = point;
    }

    private void PollAppUsage()
    {
        if (!isTracking)
        {
            lastAppUsageTickUtc = null;
            return;
        }

        var nowUtc = DateTime.UtcNow;
        var elapsedSeconds = lastAppUsageTickUtc.HasValue ? (nowUtc - lastAppUsageTickUtc.Value).TotalSeconds : 0;
        lastAppUsageTickUtc = nowUtc;

        // 系统睡眠、卡顿或暂停后产生的大跨度不计入，避免一次性把离线时间算给某个程序。
        if (elapsedSeconds <= 0 || elapsedSeconds > 10) return;

        var runningApps = RunningAppUsageSnapshot.GetRunningExternalApps()
            .Where(app => !string.IsNullOrWhiteSpace(app.Name))
            .GroupBy(app => NormalizeAppName(app.Name), StringComparer.OrdinalIgnoreCase)
            .Where(group => !IsSelfApp(group.Key))
            .Select(group => new
            {
                Name = group.Key,
                Path = group.Select(x => x.ExecutablePath).FirstOrDefault(x => !string.IsNullOrWhiteSpace(x) && File.Exists(x))
            })
            .ToArray();

        if (runningApps.Length == 0) return;

        EnsureTodayForInput();
        state.AppUsageSeconds ??= new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);
        state.AppPaths ??= new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        foreach (var app in runningApps)
        {
            state.AppUsageSeconds[app.Name] = state.AppUsageSeconds.TryGetValue(app.Name, out var oldSeconds)
                ? oldSeconds + elapsedSeconds
                : elapsedSeconds;

            if (!string.IsNullOrWhiteSpace(app.Path))
                state.AppPaths[app.Name] = app.Path;
        }

        stateDirty = true;

        // 应用详情页需要按秒实时增长，但不再清空重建真实 ListView 项，
        // 只更新虚拟数据并双缓冲重绘可见区域，避免闪烁。
        var uiSignature = string.Join("|", runningApps.Select(x => x.Name).OrderBy(x => x, StringComparer.OrdinalIgnoreCase));
        lastAppUsageUiSignature = uiSignature;
        lastAppUsageUiRefreshUtc = nowUtc;
        appDetailDirty = true;
        if (tabs.SelectedTab?.Name == "appDetailTab") liveDirty = true;
    }

    private void PollAppUsageAsync()
    {
        if (!isTracking)
        {
            lastAppUsageTickUtc = null;
            return;
        }

        var nowUtc = DateTime.UtcNow;
        var elapsedSeconds = lastAppUsageTickUtc.HasValue ? (nowUtc - lastAppUsageTickUtc.Value).TotalSeconds : 0;
        lastAppUsageTickUtc = nowUtc;

        // 系统睡眠、卡顿或暂停后产生的大跨度不计入，避免一次性把离线时间算给某个程序。
        if (elapsedSeconds <= 0 || elapsedSeconds > 10) return;

        Task.Run(() =>
        {
            try
            {
                var runningApps = RunningAppUsageSnapshot.GetRunningExternalApps()
                    .Where(app => !string.IsNullOrWhiteSpace(app.Name))
                    .GroupBy(app => NormalizeAppName(app.Name), StringComparer.OrdinalIgnoreCase)
                    .Where(group => !IsSelfApp(group.Key))
                    .Select(group => new
                    {
                        Name = group.Key,
                        Path = group.Select(x => x.ExecutablePath).FirstOrDefault(x => !string.IsNullOrWhiteSpace(x) && File.Exists(x))
                    })
                    .ToArray();

                if (runningApps.Length == 0) return;

                if (IsDisposed) return;
                BeginInvoke(new Action(() =>
                {
                    if (IsDisposed) return;

                    EnsureTodayForInput();
                    state.AppUsageSeconds ??= new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);
                    state.AppPaths ??= new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

                    foreach (var app in runningApps)
                    {
                        state.AppUsageSeconds[app.Name] = state.AppUsageSeconds.TryGetValue(app.Name, out var oldSeconds)
                            ? oldSeconds + elapsedSeconds
                            : elapsedSeconds;

                        if (!string.IsNullOrWhiteSpace(app.Path))
                            state.AppPaths[app.Name] = app.Path;
                    }

                    stateDirty = true;

                    var uiSignature = string.Join("|", runningApps.Select(x => x.Name).OrderBy(x => x, StringComparer.OrdinalIgnoreCase));
                    lastAppUsageUiSignature = uiSignature;
                    lastAppUsageUiRefreshUtc = nowUtc;
                    appDetailDirty = true;
                    if (tabs.SelectedTab?.Name == "appDetailTab") liveDirty = true;
                }));
            }
            catch { }
        });
    }

    private void CountKeyInput(string name)
    {
        if (IsMouseOnlyMode()) return;
        if (IsDisposed) return;
        if (InvokeRequired)
        {
            BeginInvoke(new Action(() => CountKeyInput(name)));
            return;
        }

        CountInput(() =>
        {
            EnsureTodayForInput();
            EnsureHourlyArrays();
            EnsureMinuteArrays();
            AddCount(state.Keys, name);
            AddAppKeyCount(name);
            keyTotalCache++;
            var nowLocal = DateTime.Now;
            recentKeyTimestamps.Enqueue(DateTime.UtcNow);
            TrimRecentKeyTimestamps(DateTime.UtcNow);
            state.HourlyKeys[nowLocal.Hour]++;
            state.MinuteKeys[nowLocal.Hour * 60 + nowLocal.Minute]++;
            TrackRapidAndCombo(name);
            UpdateOverlay(name, state.Keys.TryGetValue(name, out var count) ? count : 0);
        });
    }

    private void AddAppKeyCount(string keyName)
    {
        state.AppKeys ??= new Dictionary<string, Dictionary<string, int>>(StringComparer.OrdinalIgnoreCase);
        state.AppPaths ??= new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        if (!ForegroundApp.TryGetExternalForegroundProcess(out var pid, out var appProcessName)) return;
        var appName = NormalizeAppName(appProcessName);
        if (!state.AppKeys.TryGetValue(appName, out var keys))
        {
            keys = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            state.AppKeys[appName] = keys;
        }
        AddCount(keys, keyName);
        appDetailDirty = true;

        if (!state.AppPaths.ContainsKey(appName))
            ScheduleResolveAppPath(pid, appName);
    }

    private void ScheduleResolveAppPath(uint pid, string appName)
    {
        if (pid == 0 || string.IsNullOrWhiteSpace(appName)) return;
        if (!pendingAppPathResolves.TryAdd(pid, 0)) return;

        Task.Run(() =>
        {
            try
            {
                var path = ForegroundApp.TryGetProcessPath(pid);
                if (string.IsNullOrWhiteSpace(path) || !File.Exists(path)) return;
                if (IsDisposed) return;
                BeginInvoke(new Action(() =>
                {
                    if (IsDisposed) return;
                    state.AppPaths ??= new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                    if (!state.AppPaths.ContainsKey(appName))
                    {
                        state.AppPaths[appName] = path;
                        stateDirty = true;
                        appDetailDirty = true;
                    }
                }));
            }
            catch { }
            finally
            {
                pendingAppPathResolves.TryRemove(pid, out _);
            }
        });
    }

    private void CountInput(Action action)
    {
        if (IsDisposed) return;
        if (InvokeRequired)
        {
            BeginInvoke(new Action(() => CountInput(action)));
            return;
        }

        action();
        stateDirty = true;
        rankDirty = true;
        groupDirty = true;
        if (tabs.SelectedTab?.Name == "appDetailTab") appDetailDirty = true;
        liveDirty = true;
        // 排行、分组、应用详情统一由 100ms 定时器批量刷新，避免每次计数都重建列表造成闪烁。
    }

    private string[] GetLivePressedKeys()
    {
        return physicalDownKeys
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .Take(20)
            .ToArray();
    }

    private string[] GetLivePressedMouse()
    {
        return physicalDownMouse
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .Take(8)
            .ToArray();
    }


    private bool CleanupStalePhysicalDownKeys()
    {
        var now = DateTime.UtcNow;
        var stale = physicalDownKeys
            .Where(key => InputStateProbe.CanProbeKeyboardKey(key)
                ? !InputStateProbe.IsKeyboardKeyDown(key)
                : (!physicalDownLastSeen.TryGetValue(key, out var seen) || now - seen > PhysicalKeyStaleTimeout))
            .ToArray();

        var changed = false;
        if (stale.Length > 0)
        {
            foreach (var key in stale)
            {
                physicalDownKeys.Remove(key);
                physicalDownLastSeen.Remove(key);
            }
            changed = true;
        }

        var releasedMouse = physicalDownMouse
            .Where(button => !button.StartsWith("Wheel", StringComparison.OrdinalIgnoreCase) && !InputStateProbe.IsMouseButtonDown(button))
            .ToArray();
        foreach (var button in releasedMouse)
        {
            physicalDownMouse.Remove(button);
            mouseMomentaryUntil.Remove(button);
            changed = true;
        }

        if (mouseMomentaryUntil.Count > 0)
        {
            var expiredMouse = mouseMomentaryUntil
                .Where(x => now >= x.Value)
                .Select(x => x.Key)
                .ToArray();
            foreach (var button in expiredMouse)
            {
                mouseMomentaryUntil.Remove(button);
                physicalDownMouse.Remove(button);
                changed = true;
            }
        }

        return changed;
    }

    private void UpdateOverlay(string name, int count)
    {
        if (!overlayButton.Checked) return;
        overlayKeyName = name;
        overlayKeyCount = count;
        overlayKeyDirty = true;
    }

    private void TrimRecentKeyTimestamps(DateTime nowUtc)
    {
        while (recentKeyTimestamps.Count > 0 && nowUtc - recentKeyTimestamps.Peek() > TimeSpan.FromHours(1))
            recentKeyTimestamps.Dequeue();
    }

    private SpeedStats BuildSpeedStats(UsageState source)
    {
        // 回归正常统计口径：
        // 本分钟 = 当前自然分钟内按了多少次；
        // 当前小时 = 当前自然小时内按了多少次；
        // 平均速度 = 今日总按键 ÷ 从第一次按键到现在经过的分钟数。
        // 长按仍然只算一次。
        var minutes = NormalizeMinute(source.MinuteKeys);
        var hours = NormalizeHourly(source.HourlyKeys);
        var now = DateTime.Now;
        var minuteIndex = now.Hour * 60 + now.Minute;

        var currentMinute = minuteIndex >= 0 && minuteIndex < minutes.Length ? minutes[minuteIndex] : 0;
        var currentHour = now.Hour >= 0 && now.Hour < hours.Length ? hours[now.Hour] : 0;

        var maxMinute = minutes.Length == 0 ? 0 : minutes.Max();
        var maxHour = hours.Length == 0 ? 0 : hours.Max();
        var activeAverage = minutes.Where(v => v > 0).DefaultIfEmpty(0).Average();

        var total = source.Keys.Values.Sum();
        var firstMinuteIndex = Array.FindIndex(minutes, v => v > 0);
        var elapsedMinutes = firstMinuteIndex >= 0 ? Math.Max(1, minuteIndex - firstMinuteIndex + 1) : 0;
        var averageMinute = elapsedMinutes > 0 ? total / (double)elapsedMinutes : 0;
        var averageHour = averageMinute * 60.0;

        return new SpeedStats(currentMinute, currentHour, maxMinute, maxHour, activeAverage, averageMinute, averageHour);
    }

    private void ConfigureTray()
    {
        trayIcon.Text = "KeyMouse Heatmap";
        trayIcon.Icon = appIcon;
        trayIcon.Visible = false;
        trayIcon.DoubleClick += (_, _) => RestoreFromTray();
        trayIcon.ContextMenuStrip = new ContextMenuStrip();
        trayIcon.ContextMenuStrip.Items.Add("显示", null, (_, _) => RestoreFromTray());
        trayIcon.ContextMenuStrip.Items.Add("开始/暂停", null, (_, _) => ToggleTracking());
        trayIcon.ContextMenuStrip.Items.Add("清空今日", null, (_, _) => ResetCurrentDay());
        trayIcon.ContextMenuStrip.Items.Add("打开数据目录", null, (_, _) => Process.Start(new ProcessStartInfo("explorer.exe", appDir) { UseShellExecute = true }));
        trayIcon.ContextMenuStrip.Items.Add("退出", null, (_, _) => Close());
    }

    private void RestoreFromTray()
    {
        Show();
        WindowState = FormWindowState.Normal;
        Activate();
    }

    private void ToggleTracking()
    {
        if (isTracking) StopTracking();
        else StartTracking();
    }

    private void StartTracking()
    {
        ChangeDate(DateOnly.FromDateTime(DateTime.Today));
        lastMousePoint = null;
        lastAppUsageTickUtc = DateTime.UtcNow;
        physicalDownKeys.Clear();
        physicalDownLastSeen.Clear();
        physicalDownMouse.Clear();
        mouseMomentaryUntil.Clear();
        RawKeyboardInput.Register(Handle);
        RawMouseInput.Register(Handle);
        mouseDistanceTimer.Start();
        appUsageTimer.Start();
        isTracking = true;
        RefreshAll();
    }

    private void StopTracking()
    {
        PollAppUsage();
        RawKeyboardInput.Unregister();
        RawMouseInput.Unregister();
        mouseDistanceTimer.Stop();
        appUsageTimer.Stop();
        lastMousePoint = null;
        lastAppUsageTickUtc = null;
        physicalDownKeys.Clear();
        physicalDownLastSeen.Clear();
        physicalDownMouse.Clear();
        mouseMomentaryUntil.Clear();
        pressedDirty = true;
        isTracking = false;
        SaveState(force: true);
        RefreshAll();
    }

    private void ChangeDate(DateOnly date)
    {
        SaveState(force: true);
        selectedDate = date;
        state = LoadState(selectedDate);
        lastMousePoint = null;
        RecalculateTotals();
        if (DateOnly.FromDateTime(datePicker.Value) != selectedDate)
            datePicker.Value = selectedDate.ToDateTime(TimeOnly.MinValue);
        combinedView.Invalidate();
        keyboardView.Invalidate();
        mouseView.Invalidate();
        peakView.Invalidate();
        speedView.Invalidate();
        rankDirty = true;
        groupDirty = true;
        appDetailDirty = true;
        RefreshAll();
    }

    private void EnsureTodayForInput()
    {
        var today = DateOnly.FromDateTime(DateTime.Today);
        if (selectedDate == today) return;
        SaveState(force: true);
        selectedDate = today;
        state = LoadState(selectedDate);
        RecalculateTotals();
        datePicker.Value = selectedDate.ToDateTime(TimeOnly.MinValue);
        rankDirty = true;
        groupDirty = true;
        appDetailDirty = true;
    }

    private void RefreshAll()
    {
        displayState = BuildDisplayState();
        rankDirty = true;
        groupDirty = true;
        appDetailDirty = true;
        RefreshLive();
        RefreshRankList();
        RefreshGroupStats();
        RefreshAppDetailList();
    }

    private void RefreshLive()
    {
        if (IsHeavyUiSuspended()) return;
        if (IsUiActive())
        {
            startButton.Text = isTracking ? "暂停" : "开始";
            nextButton.Enabled = selectedDate < DateOnly.FromDateTime(DateTime.Today);
            stateLabel.Text = $"{AppVersion} | 状态: {(isTracking ? "记录中" : "已暂停")} | 日期: {selectedDate:yyyy-MM-dd}";
            totalLabel.Text = $"今日 键盘 {keyTotalCache} 次 | 鼠标 {mouseTotalCache} 次 | 显示范围 {rangeBox.SelectedItem ?? "今天"}";
            totalLabel.Text += $" | \u9f20\u6807\u79fb\u52a8 {FormatDistance(state.MouseDistancePixels)}";
            pathLabel.Text = CurrentDataFile;
            InvalidateActiveView();
        }
        if (overlayButton.Checked) overlay.UpdateSpeed(BuildSpeedStats(state));
    }

    private bool IsUiActive() => Visible && WindowState != FormWindowState.Minimized && !IsDisposed;

    private void InvalidateActiveView(bool immediate = false)
    {
        var selected = tabs.SelectedTab?.Controls.Cast<Control>().FirstOrDefault();
        if (selected is ListView) return;
        if (selected == null || selected.IsDisposed) return;

        selected.Invalidate();
        // 键盘/鼠标单页很轻，可以同步 Update 保证即时高亮。
        // 总览页包含键盘、鼠标和 Top10 卡片，同步 Update 会在 Raw Input 消息里强制重绘整页，
        // v1.7.4 的卡顿主要就是这里把输入消息堵住了，所以总览页改为异步合并重绘。
        if (!immediate || !selected.IsHandleCreated || !selected.Visible) return;

        selected.Update();
    }

    private void DrawAppNameCell(Graphics g, Rectangle bounds, AppKeyRow row, bool selected)
    {
        var bgColor = selected ? Color.FromArgb(219, 235, 255) : appDetailList.BackColor;
        using var bg = new SolidBrush(bgColor);
        g.FillRectangle(bg, bounds);

        var iconKey = appIconList.Images.ContainsKey(row.IconKey) ? row.IconKey : "__default";
        var image = appIconList.Images[iconKey] ?? appIconList.Images["__default"];
        var iconSize = Math.Min(20, Math.Max(16, bounds.Height - 6));
        var iconRect = new Rectangle(bounds.X + 6, bounds.Y + (bounds.Height - iconSize) / 2, iconSize, iconSize);
        if (image != null) g.DrawImage(image, iconRect);

        var textRect = new Rectangle(iconRect.Right + 8, bounds.Y, Math.Max(0, bounds.Width - iconSize - 18), bounds.Height);
        TextRenderer.DrawText(g, row.App, appDetailList.Font, textRect, appDetailList.ForeColor, TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis);
    }

    private static void DrawStatBar(Graphics g, Rectangle bounds, int value, int max)
    {
        var outer = Rectangle.Inflate(bounds, -8, -7);
        if (outer.Width <= 8 || outer.Height <= 6) return;

        using var back = new SolidBrush(Color.FromArgb(236, 240, 245));
        using var fill = new SolidBrush(Color.FromArgb(55, 132, 245));
        using var border = new Pen(Color.FromArgb(197, 208, 222));
        g.FillRectangle(back, outer);
        g.DrawRectangle(border, outer);

        var ratio = max <= 0 ? 0 : Math.Clamp(value / (double)max, 0, 1);
        var fillRect = new Rectangle(outer.X + 1, outer.Y + 1, Math.Max(0, (int)Math.Round((outer.Width - 2) * ratio)), outer.Height - 2);
        g.FillRectangle(fill, fillRect);

        var text = max <= 0 ? "" : $"{ratio:P1}";
        TextRenderer.DrawText(g, text, SystemFonts.MessageBoxFont, outer, Color.FromArgb(28, 43, 58), TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis);
    }

    private void RefreshRankList()
    {
        if (tabs.SelectedTab?.Text != "排行") return;
        if (!rankDirty) return;
        if (IsHeavyUiSuspended()) return;

        rankDirty = false;
        displayState = BuildDisplayState();
        var rows = displayState.Keys.Select(x => new RankRow(0, "键盘", x.Key, x.Value))
            .Concat(displayState.Mouse.Select(x => new RankRow(0, "鼠标", x.Key, x.Value)))
            .Where(x => x.Count > 0);
        rankRows = SortRankRows(rows)
            .Take(200)
            .Select((x, index) => x with { Rank = index + 1 })
            .ToList();
        rankTotalCount = Math.Max(1, rankRows.Sum(x => x.Count));
        UpdateRankColumnHeaders();
        FitRankColumns();

        rankList.BeginUpdate();
        if (rankList.VirtualListSize != rankRows.Count)
            rankList.VirtualListSize = rankRows.Count;
        rankList.Invalidate();
        rankList.EndUpdate();
    }

    private void RefreshAppDetailList()
    {
        if (tabs.SelectedTab?.Name != "appDetailTab") return;
        if (!appDetailDirty) return;
        if (IsHeavyUiSuspended()) return;

        appDetailDirty = false;
        displayState = BuildDisplayState();
        appKeyRows = SortAppRows(BuildAppSummaryRows(displayState))
            .Take(300)
            .ToList();
        appTotalCount = Math.Max(1, appKeyRows.Sum(x => x.Count));
        EnsureAppIcons(appKeyRows);
        UpdateAppColumnHeaders();
        FitAppColumns();

        appDetailList.BeginUpdate();
        if (appDetailList.VirtualListSize != appKeyRows.Count)
            appDetailList.VirtualListSize = appKeyRows.Count;
        appDetailList.Invalidate();
        appDetailList.EndUpdate();
    }

    private IEnumerable<RankRow> SortRankRows(IEnumerable<RankRow> rows)
    {
        IOrderedEnumerable<RankRow> ordered = rankSortColumn switch
        {
            0 => rankSortOrder == SortOrder.Ascending ? rows.OrderBy(x => x.Count).ThenBy(x => x.Device).ThenBy(x => x.Name) : rows.OrderByDescending(x => x.Count).ThenBy(x => x.Device).ThenBy(x => x.Name),
            1 => rankSortOrder == SortOrder.Ascending ? rows.OrderBy(x => x.Device).ThenByDescending(x => x.Count).ThenBy(x => x.Name) : rows.OrderByDescending(x => x.Device).ThenByDescending(x => x.Count).ThenBy(x => x.Name),
            2 => rankSortOrder == SortOrder.Ascending ? rows.OrderBy(x => x.Name).ThenBy(x => x.Device) : rows.OrderByDescending(x => x.Name).ThenBy(x => x.Device),
            3 or 4 => rankSortOrder == SortOrder.Ascending ? rows.OrderBy(x => x.Count).ThenBy(x => x.Device).ThenBy(x => x.Name) : rows.OrderByDescending(x => x.Count).ThenBy(x => x.Device).ThenBy(x => x.Name),
            _ => rows.OrderByDescending(x => x.Count).ThenBy(x => x.Device).ThenBy(x => x.Name)
        };
        return ordered;
    }

    private IEnumerable<(string Group, int Count, string Note)> SortGroupRows(IEnumerable<(string Group, int Count, string Note)> rows)
    {
        IOrderedEnumerable<(string Group, int Count, string Note)> ordered = groupSortColumn switch
        {
            0 => groupSortOrder == SortOrder.Ascending ? rows.OrderBy(x => x.Group) : rows.OrderByDescending(x => x.Group),
            1 or 2 => groupSortOrder == SortOrder.Ascending ? rows.OrderBy(x => x.Count).ThenBy(x => x.Group) : rows.OrderByDescending(x => x.Count).ThenBy(x => x.Group),
            3 => groupSortOrder == SortOrder.Ascending ? rows.OrderBy(x => x.Note).ThenBy(x => x.Group) : rows.OrderByDescending(x => x.Note).ThenBy(x => x.Group),
            _ => rows.OrderByDescending(x => x.Count).ThenBy(x => x.Group)
        };
        return ordered;
    }

    private IEnumerable<AppKeyRow> SortAppRows(IEnumerable<AppKeyRow> rows)
    {
        IOrderedEnumerable<AppKeyRow> ordered = appSortColumn switch
        {
            0 => appSortOrder == SortOrder.Ascending ? rows.OrderBy(x => x.App) : rows.OrderByDescending(x => x.App),
            1 => appSortOrder == SortOrder.Ascending ? rows.OrderBy(x => x.UsageSeconds).ThenBy(x => x.App) : rows.OrderByDescending(x => x.UsageSeconds).ThenByDescending(x => x.Count).ThenBy(x => x.App),
            2 or 4 => appSortOrder == SortOrder.Ascending ? rows.OrderBy(x => x.Count).ThenBy(x => x.App) : rows.OrderByDescending(x => x.Count).ThenBy(x => x.App),
            3 => appSortOrder == SortOrder.Ascending ? rows.OrderBy(x => x.KeyTypes).ThenBy(x => x.App) : rows.OrderByDescending(x => x.KeyTypes).ThenBy(x => x.App),
            5 => appSortOrder == SortOrder.Ascending ? rows.OrderBy(x => x.Summary).ThenBy(x => x.App) : rows.OrderByDescending(x => x.Summary).ThenBy(x => x.App),
            _ => rows.OrderByDescending(x => x.UsageSeconds).ThenByDescending(x => x.Count).ThenBy(x => x.App)
        };
        return ordered;
    }

    private void ToggleRankSort(int column)
    {
        if (rankSortColumn == column) rankSortOrder = ToggleSortOrder(rankSortOrder);
        else
        {
            rankSortColumn = column;
            rankSortOrder = column is 0 or 3 or 4 ? SortOrder.Descending : SortOrder.Ascending;
        }
        CaptureAndQueueSettingsSave();
    }

    private void ToggleGroupSort(int column)
    {
        if (groupSortColumn == column) groupSortOrder = ToggleSortOrder(groupSortOrder);
        else
        {
            groupSortColumn = column;
            groupSortOrder = column is 1 or 2 ? SortOrder.Descending : SortOrder.Ascending;
        }
        CaptureAndQueueSettingsSave();
    }

    private void ToggleAppSort(int column)
    {
        if (appSortColumn == column) appSortOrder = ToggleSortOrder(appSortOrder);
        else
        {
            appSortColumn = column;
            appSortOrder = column is 1 or 2 or 3 or 4 ? SortOrder.Descending : SortOrder.Ascending;
        }
        combinedView.AppSortColumn = appSortColumn;
        combinedView.AppSortOrder = appSortOrder;
        CaptureAndQueueSettingsSave();
    }

    private static SortOrder ToggleSortOrder(SortOrder order) => order == SortOrder.Ascending ? SortOrder.Descending : SortOrder.Ascending;

    private static SortOrder NormalizeSortOrder(SortOrder value, SortOrder fallback) =>
        value is SortOrder.Ascending or SortOrder.Descending ? value : fallback;

    private static string Header(string title, int column, int sortColumn, SortOrder order)
    {
        if (column != sortColumn) return title;
        return title + (order == SortOrder.Ascending ? " ▲" : " ▼");
    }

    private void UpdateRankColumnHeaders()
    {
        SetColumnText(rankList, 0, Header("排行", 0, rankSortColumn, rankSortOrder));
        SetColumnText(rankList, 1, Header("设备", 1, rankSortColumn, rankSortOrder));
        SetColumnText(rankList, 2, Header("按键", 2, rankSortColumn, rankSortOrder));
        SetColumnText(rankList, 3, Header("次数", 3, rankSortColumn, rankSortOrder));
        SetColumnText(rankList, 4, Header("统计条", 4, rankSortColumn, rankSortOrder));
    }

    private void UpdateGroupColumnHeaders()
    {
        SetColumnText(groupList, 0, Header("分组", 0, groupSortColumn, groupSortOrder));
        SetColumnText(groupList, 1, Header("次数", 1, groupSortColumn, groupSortOrder));
        SetColumnText(groupList, 2, Header("统计条", 2, groupSortColumn, groupSortOrder));
        SetColumnText(groupList, 3, Header("说明", 3, groupSortColumn, groupSortOrder));
    }

    private void UpdateAppColumnHeaders()
    {
        SetColumnText(appDetailList, 0, Header("应用", 0, appSortColumn, appSortOrder));
        SetColumnText(appDetailList, 1, Header("使用时长", 1, appSortColumn, appSortOrder));
        SetColumnText(appDetailList, 2, Header("总次数", 2, appSortColumn, appSortOrder));
        SetColumnText(appDetailList, 3, Header("按键种类", 3, appSortColumn, appSortOrder));
        SetColumnText(appDetailList, 4, Header("统计条", 4, appSortColumn, appSortOrder));
        SetColumnText(appDetailList, 5, Header("常用按键", 5, appSortColumn, appSortOrder));
    }

    private static void SetColumnText(ListView list, int index, string text)
    {
        if (list.Columns.Count > index) list.Columns[index].Text = text;
    }

    private void FitRankColumns()
    {
        if (rankList.Columns.Count < 5 || rankList.ClientSize.Width <= 0) return;
        var w = Math.Max(520, rankList.ClientSize.Width - SystemInformation.VerticalScrollBarWidth - 6);
        rankList.Columns[0].Width = Math.Clamp((int)(w * 0.12), 58, 90);
        rankList.Columns[1].Width = Math.Clamp((int)(w * 0.16), 80, 130);
        rankList.Columns[2].Width = Math.Clamp((int)(w * 0.30), 150, 260);
        rankList.Columns[3].Width = Math.Clamp((int)(w * 0.16), 90, 130);
        rankList.Columns[4].Width = Math.Max(170, w - rankList.Columns[0].Width - rankList.Columns[1].Width - rankList.Columns[2].Width - rankList.Columns[3].Width);
    }

    private void FitGroupColumns()
    {
        if (groupList.Columns.Count < 4 || groupList.ClientSize.Width <= 0) return;
        var w = Math.Max(620, groupList.ClientSize.Width - SystemInformation.VerticalScrollBarWidth - 6);
        groupList.Columns[0].Width = Math.Clamp((int)(w * 0.22), 130, 220);
        groupList.Columns[1].Width = Math.Clamp((int)(w * 0.14), 90, 140);
        groupList.Columns[2].Width = Math.Clamp((int)(w * 0.26), 170, 280);
        groupList.Columns[3].Width = Math.Max(240, w - groupList.Columns[0].Width - groupList.Columns[1].Width - groupList.Columns[2].Width);
    }

    private void FitAppColumns()
    {
        if (appDetailList.Columns.Count < 6 || appDetailList.ClientSize.Width <= 0) return;
        var w = Math.Max(760, appDetailList.ClientSize.Width - SystemInformation.VerticalScrollBarWidth - 6);
        appDetailList.Columns[0].Width = Math.Clamp((int)(w * 0.22), 170, 260);
        appDetailList.Columns[1].Width = Math.Clamp((int)(w * 0.15), 110, 165);
        appDetailList.Columns[2].Width = Math.Clamp((int)(w * 0.12), 90, 130);
        appDetailList.Columns[3].Width = Math.Clamp((int)(w * 0.12), 90, 130);
        appDetailList.Columns[4].Width = Math.Clamp((int)(w * 0.20), 160, 260);
        appDetailList.Columns[5].Width = Math.Max(260, w - appDetailList.Columns[0].Width - appDetailList.Columns[1].Width - appDetailList.Columns[2].Width - appDetailList.Columns[3].Width - appDetailList.Columns[4].Width);
    }

    private void ListViewZoomMouseWheel(object? sender, MouseEventArgs e)
    {
        if ((ModifierKeys & Keys.Control) != Keys.Control) return;
        listZoom = Math.Clamp(listZoom + (e.Delta > 0 ? 0.08f : -0.08f), 0.82f, 1.55f);
        ApplyListZoom();
    }

    private void ApplyListZoom()
    {
        var size = Math.Clamp(baseListFont.Size * listZoom, 8.0f, 15.0f);
        rankList.Font = new Font(baseListFont.FontFamily, size, baseListFont.Style);
        groupList.Font = new Font(baseListFont.FontFamily, size, baseListFont.Style);
        appDetailList.Font = new Font(baseListFont.FontFamily, size, baseListFont.Style);
        FitRankColumns();
        FitGroupColumns();
        FitAppColumns();
        rankList.Invalidate();
        groupList.Invalidate();
        appDetailList.Invalidate();
        CaptureAndQueueSettingsSave();
    }

    private bool IsSelfApp(string appName) =>
        NormalizeAppName(appName).StartsWith(selfProcessName, StringComparison.OrdinalIgnoreCase);

    private IEnumerable<AppKeyRow> BuildAppSummaryRows(UsageState source)
    {
        var appKeys = source.AppKeys ?? new Dictionary<string, Dictionary<string, int>>(StringComparer.OrdinalIgnoreCase);
        var appUsage = source.AppUsageSeconds ?? new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);

        return appKeys.Keys
            .Concat(appUsage.Keys)
            .Where(app => !IsSelfApp(app))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .GroupBy(app => NormalizeAppName(app), StringComparer.OrdinalIgnoreCase)
            .Select(group =>
            {
                var merged = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                var usageSeconds = 0.0;

                foreach (var appName in group)
                {
                    if (appKeys.TryGetValue(appName, out var keys))
                        MergeCounts(merged, keys);
                    if (appUsage.TryGetValue(appName, out var seconds))
                        usageSeconds += seconds;
                }

                var top = merged
                    .Where(x => x.Value > 0)
                    .OrderByDescending(x => x.Value)
                    .ThenBy(x => x.Key)
                    .Take(50)
                    .Select(x => $"{x.Key} {x.Value}")
                    .ToArray();
                var iconPath = group
                    .Select(x => TryGetAppPath(source, x))
                    .FirstOrDefault(x => !string.IsNullOrWhiteSpace(x));
                iconPath ??= AppIconProvider.FindRunningProcessPath(group.Key);
                var summary = top.Length > 0 ? string.Join(", ", top) : "暂无按键记录";
                return new AppKeyRow(group.Key, summary, merged.Values.Sum(), merged.Count(x => x.Value > 0), usageSeconds, AppIconProvider.IconKey(group.Key, iconPath), iconPath);
            })
            .Where(x => x.Count > 0 || x.UsageSeconds >= 1);
    }

    private static string? TryGetAppPath(UsageState source, string appName)
    {
        if (source.AppPaths != null && source.AppPaths.TryGetValue(appName, out var path) && File.Exists(path)) return path;
        var normalized = NormalizeAppName(appName);
        return source.AppPaths?
            .Where(x => string.Equals(NormalizeAppName(x.Key), normalized, StringComparison.OrdinalIgnoreCase) && File.Exists(x.Value))
            .Select(x => x.Value)
            .FirstOrDefault();
    }

    private void EnsureAppIcons(IEnumerable<AppKeyRow> rows)
    {
        foreach (var row in rows)
        {
            if (appIconList.Images.ContainsKey(row.IconKey)) continue;
            appIconList.Images.Add(row.IconKey, AppIconProvider.GetIconBitmap(row.IconPath));
        }
    }

    private static string NormalizeAppName(string appName)
    {
        if (string.IsNullOrWhiteSpace(appName)) return "Unknown";
        var trimmed = appName.Trim();
        var dash = trimmed.IndexOf(" - ", StringComparison.Ordinal);
        if (dash > 0) trimmed = trimmed[..dash].Trim();
        if (trimmed.EndsWith(".exe", StringComparison.OrdinalIgnoreCase)) trimmed = trimmed[..^4];
        return string.IsNullOrWhiteSpace(trimmed) ? "Unknown" : trimmed;
    }

    private void ResetCurrentDay()
    {
        var confirm = MessageBox.Show($"确定清空 {selectedDate:yyyy-MM-dd} 的统计吗？", "确认清空", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
        if (confirm != DialogResult.Yes) return;
        state = NewState(selectedDate);
        RecalculateTotals();
        SaveState(force: true);
        rankDirty = true;
        groupDirty = true;
        appDetailDirty = true;
        RefreshAll();
    }

    private void ExportCsv()
    {
        SaveState();
        var csvPath = Path.Combine(appDir, $"usage-{selectedDate:yyyy-MM-dd}.csv");
        using var writer = new StreamWriter(csvPath, false, new UTF8Encoding(true));
        writer.WriteLine("Date,Device,Button,Count");
        foreach (var entry in state.Keys.OrderByDescending(x => x.Value)) writer.WriteLine($"{selectedDate:yyyy-MM-dd},Keyboard,\"{EscapeCsv(entry.Key)}\",{entry.Value}");
        foreach (var entry in state.Mouse.OrderByDescending(x => x.Value)) writer.WriteLine($"{selectedDate:yyyy-MM-dd},Mouse,\"{EscapeCsv(entry.Key)}\",{entry.Value}");
        writer.WriteLine();
        writer.WriteLine("Date,App,UsageSeconds,UsageTime,TotalCount,KeyTypes,TopKeys");
        foreach (var app in BuildAppSummaryRows(state).OrderByDescending(x => x.UsageSeconds).ThenByDescending(x => x.Count).ThenBy(x => x.App))
            writer.WriteLine($"{selectedDate:yyyy-MM-dd},\"{EscapeCsv(app.App)}\",{app.UsageSeconds:F0},\"{EscapeCsv(FormatDuration(app.UsageSeconds))}\",{app.Count},{app.KeyTypes},\"{EscapeCsv(app.Summary)}\"");
        writer.WriteLine();
        writer.WriteLine("Date,Metric,Value");
        writer.WriteLine($"{selectedDate:yyyy-MM-dd},MouseDistancePixels,{state.MouseDistancePixels:F0}");
        MessageBox.Show($"已导出到：\n{csvPath}", "导出完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }

    private void ExportPng()
    {
        SaveState();
        using var dialog = new SaveFileDialog
        {
            Filter = "PNG 图片|*.png",
            FileName = $"heatmap-{selectedDate:yyyy-MM-dd}.png",
            InitialDirectory = appDir
        };
        if (dialog.ShowDialog(this) != DialogResult.OK) return;

        using var bitmap = new Bitmap(1200, 780);
        using var g = Graphics.FromImage(bitmap);
        g.Clear(Color.FromArgb(247, 248, 250));
        HeatmapRenderer.DrawKeyboard(g, new Rectangle(24, 24, 1152, 420), state, true, Array.Empty<string>());
        HeatmapRenderer.DrawMouse(g, new Rectangle(24, 470, 1152, 280), state);
        bitmap.Save(dialog.FileName, System.Drawing.Imaging.ImageFormat.Png);
        MessageBox.Show($"已导出到：\n{dialog.FileName}", "导出完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }


    private UsageState BuildDisplayState()
    {
        var range = rangeBox.SelectedItem?.ToString() ?? "今天";
        if (range == "今天") return state;

        var total = NewState(selectedDate);
        var minDate = range switch
        {
            "最近7天" => DateOnly.FromDateTime(DateTime.Today.AddDays(-6)),
            "最近30天" => DateOnly.FromDateTime(DateTime.Today.AddDays(-29)),
            _ => (DateOnly?)null
        };

        foreach (var file in Directory.GetFiles(appDir, "usage-*.json"))
        {
            var name = Path.GetFileNameWithoutExtension(file).Replace("usage-", "");
            if (!DateOnly.TryParse(name, out var day)) continue;
            if (minDate.HasValue && day < minDate.Value) continue;
            var one = day == selectedDate ? state : LoadState(day);
            MergeCounts(total.Keys, one.Keys);
            MergeCounts(total.Mouse, one.Mouse);
            MergeCounts(total.Combos, one.Combos);
            MergeCounts(total.RapidKeys, one.RapidKeys);
            MergeAppKeys(total.AppKeys, one.AppKeys);
            MergeAppUsage(total.AppUsageSeconds, one.AppUsageSeconds);
            MergeAppPaths(total.AppPaths, one.AppPaths);
            total.MouseDistancePixels += one.MouseDistancePixels;
            var hk = NormalizeHourly(one.HourlyKeys);
            var hm = NormalizeHourly(one.HourlyMouse);
            var mk = NormalizeMinute(one.MinuteKeys);
            for (var i = 0; i < 1440; i++) total.MinuteKeys[i] += mk[i];
            for (var i = 0; i < 24; i++)
            {
                total.HourlyKeys[i] += hk[i];
                total.HourlyMouse[i] += hm[i];
            }
        }
        total.Date = range;
        return total;
    }

    private static void MergeCounts(Dictionary<string, int> target, Dictionary<string, int>? source)
    {
        if (source == null) return;
        foreach (var pair in source)
            target[pair.Key] = target.TryGetValue(pair.Key, out var old) ? old + pair.Value : pair.Value;
    }

    private static void MergeAppKeys(Dictionary<string, Dictionary<string, int>> target, Dictionary<string, Dictionary<string, int>>? source)
    {
        if (source == null) return;
        foreach (var app in source)
        {
            if (!target.TryGetValue(app.Key, out var keys))
            {
                keys = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                target[app.Key] = keys;
            }
            MergeCounts(keys, app.Value);
        }
    }

    private static void MergeAppUsage(Dictionary<string, double> target, Dictionary<string, double>? source)
    {
        if (source == null) return;
        foreach (var pair in source)
            target[pair.Key] = target.TryGetValue(pair.Key, out var old) ? old + pair.Value : pair.Value;
    }

    private static void MergeAppPaths(Dictionary<string, string> target, Dictionary<string, string>? source)
    {
        if (source == null) return;
        foreach (var pair in source)
        {
            if (!target.ContainsKey(pair.Key) && !string.IsNullOrWhiteSpace(pair.Value))
                target[pair.Key] = pair.Value;
        }
    }

    private readonly Dictionary<string, DateTime> lastTapAt = new(StringComparer.OrdinalIgnoreCase);

    private void TrackRapidAndCombo(string name)
    {
        state.RapidKeys ??= new Dictionary<string, int>();
        state.Combos ??= new Dictionary<string, int>();

        var now = DateTime.UtcNow;
        if (lastTapAt.TryGetValue(name, out var last) && now - last <= TimeSpan.FromSeconds(1))
            AddCount(state.RapidKeys, name);
        lastTapAt[name] = now;

        var comboKeys = physicalDownKeys.Append(name)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(x => x)
            .Take(6)
            .ToArray();
        if (comboKeys.Length >= 2)
            AddCount(state.Combos, string.Join(" + ", comboKeys));
    }


    private void RefreshGroupStats()
    {
        if (tabs.SelectedTab?.Text != "分组") return;
        if (!groupDirty) return;
        if (IsHeavyUiSuspended()) return;

        groupDirty = false;
        displayState = BuildDisplayState();
        groupRows = SortGroupRows(BuildGroupRows(displayState)).ToList();
        groupTotalCount = Math.Max(1, groupRows.Sum(x => x.Count));
        UpdateGroupColumnHeaders();
        FitGroupColumns();

        groupList.BeginUpdate();
        if (groupList.VirtualListSize != groupRows.Count)
            groupList.VirtualListSize = groupRows.Count;
        groupList.Invalidate();
        groupList.EndUpdate();
    }

    private void ExportHtmlReport()
    {
        SaveState();
        displayState = BuildDisplayState();

        var htmlPath = Path.Combine(appDir, $"report-{DateTime.Now:yyyyMMdd-HHmmss}.html");
        var rows = displayState.Keys.Select(x => new RankRow(0, "键盘", x.Key, x.Value))
            .Concat(displayState.Mouse.Select(x => new RankRow(0, "鼠标", x.Key, x.Value)))
            .Where(x => x.Count > 0)
            .OrderByDescending(x => x.Count)
            .ThenBy(x => x.Name)
            .Take(100)
            .Select((x, i) => x with { Rank = i + 1 })
            .ToList();
        var appRows = BuildAppSummaryRows(displayState)
            .OrderByDescending(x => x.UsageSeconds)
            .ThenByDescending(x => x.Count)
            .ThenBy(x => x.App)
            .Take(200)
            .ToList();

        var sb = new StringBuilder();
        sb.AppendLine("<!doctype html><meta charset='utf-8'><title>KeyMouse Heatmap Report</title>");
        sb.AppendLine("<style>body{font-family:Segoe UI,Microsoft YaHei,sans-serif;background:#f6f8fb;color:#1f2937;padding:28px}.card{background:white;border-radius:14px;padding:18px;margin:16px 0;box-shadow:0 4px 18px #0001}table{border-collapse:collapse;background:white;width:100%}td,th{border:1px solid #d8dee8;padding:8px 12px;text-align:left}th{background:#eef3f8}.num{text-align:right}</style>");
        sb.AppendLine($"<h1>KeyMouse Heatmap 报告</h1><p>范围：{Html(rangeBox.SelectedItem?.ToString() ?? "今天")}　导出时间：{DateTime.Now:yyyy-MM-dd HH:mm:ss}</p>");
        sb.AppendLine($"<div class='card'><h2>总览</h2><p>键盘：{displayState.Keys.Values.Sum()} 次；鼠标：{displayState.Mouse.Values.Sum()} 次；鼠标移动：{FormatDistance(displayState.MouseDistancePixels)}；应用使用时长：{FormatDuration(displayState.AppUsageSeconds?.Values.Sum() ?? 0)}；组合键：{displayState.Combos?.Values.Sum() ?? 0} 次；连击：{displayState.RapidKeys?.Values.Sum() ?? 0} 次。</p></div>");
        sb.AppendLine("<div class='card'><h2>分组统计</h2><table><tr><th>分组</th><th class='num'>次数</th><th>说明</th></tr>");
        foreach (var row in BuildGroupRows(displayState))
            sb.AppendLine($"<tr><td>{Html(row.Group)}</td><td class='num'>{row.Count}</td><td>{Html(row.Note)}</td></tr>");
        sb.AppendLine("</table></div>");
        sb.AppendLine("<div class='card'><h2>排行 Top 100</h2><table><tr><th>排行</th><th>设备</th><th>按键</th><th class='num'>次数</th></tr>");
        foreach (var row in rows)
            sb.AppendLine($"<tr><td>{row.Rank}</td><td>{Html(row.Device)}</td><td>{Html(row.Name)}</td><td class='num'>{row.Count}</td></tr>");
        sb.AppendLine("</table></div>");
        sb.AppendLine("<div class='card'><h2>\u5e94\u7528\u8be6\u60c5 Top 200</h2><table><tr><th>\u5e94\u7528</th><th class='num'>使用时长</th><th class='num'>\u603b\u6b21\u6570</th><th class='num'>\u6309\u952e\u79cd\u7c7b</th><th>\u5e38\u7528\u6309\u952e</th></tr>");
        foreach (var row in appRows)
            sb.AppendLine($"<tr><td>{Html(row.App)}</td><td class='num'>{Html(FormatDuration(row.UsageSeconds))}</td><td class='num'>{row.Count}</td><td class='num'>{row.KeyTypes}</td><td>{Html(row.Summary)}</td></tr>");
        sb.AppendLine("</table></div>");

        File.WriteAllText(htmlPath, sb.ToString(), Encoding.UTF8);
        MessageBox.Show($"已导出 HTML 报告：\n{htmlPath}", "导出完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }

    private static string Html(string value) => System.Net.WebUtility.HtmlEncode(value);

    private static IEnumerable<(string Group, int Count, string Note)> BuildGroupRows(UsageState s)
    {
        int Sum(params string[] names) => names.Sum(n => s.Keys.TryGetValue(n, out var v) ? v : 0);
        var letters = s.Keys.Where(kv => kv.Key.Length == 1 && kv.Key[0] >= 'A' && kv.Key[0] <= 'Z').Sum(kv => kv.Value);
        var nums = s.Keys.Where(kv => kv.Key.Length == 1 && kv.Key[0] >= '0' && kv.Key[0] <= '9').Sum(kv => kv.Value);
        var fn = s.Keys.Where(kv => kv.Key.StartsWith("F") && int.TryParse(kv.Key[1..], out _)).Sum(kv => kv.Value);
        var numpad = s.Keys.Where(kv => kv.Key.StartsWith("Num ")).Sum(kv => kv.Value);
        yield return ("字母区", letters, "A-Z 主键区字母");
        yield return ("主键数字区", nums, "键盘上方 0-9");
        yield return ("功能键区", fn, "F1-F12");
        yield return ("方向键区", Sum("Up", "Down", "Left", "Right"), "上、下、左、右");
        yield return ("小数字键盘", numpad, "Num 0-9、Num Enter、Num 运算符");
        yield return ("修饰键", Sum("Left Shift", "Right Shift", "Left Ctrl", "Right Ctrl", "Left Alt", "Right Alt", "Left Win", "Right Win", "Menu"), "左右 Shift/Ctrl/Alt/Win/Menu 分开统计");
        yield return ("鼠标", s.Mouse.Values.Sum(), "左键、右键、中键、侧键、滚轮");
        yield return ("组合键", s.Combos?.Values.Sum() ?? 0, "同时按下两个及以上键");
        yield return ("连击", s.RapidKeys?.Values.Sum() ?? 0, "1 秒内重复点击同一个键");
    }

    private string SettingsFile => Path.Combine(appDir, "settings.json");

    private AppSettings LoadSettings()
    {
        try
        {
            if (File.Exists(SettingsFile))
                return JsonSerializer.Deserialize<AppSettings>(File.ReadAllText(SettingsFile, Encoding.UTF8)) ?? new AppSettings();
        }
        catch { }
        return new AppSettings();
    }

    private void SaveSettings()
    {
        try
        {
            settingsDirty = false;
            Directory.CreateDirectory(appDir);
            File.WriteAllText(SettingsFile, JsonSerializer.Serialize(settings, jsonOptions), Encoding.UTF8);
        }
        catch { }
    }

    private void CaptureAndQueueSettingsSave()
    {
        CaptureWindowSettings();
        settingsDirty = true;
        settingsSaveTimer.Stop();
        settingsSaveTimer.Start();
    }

    private void ApplySavedWindowSettings()
    {
        if (settings.Width >= 900 && settings.Height >= 500) Size = new Size(settings.Width, settings.Height);
        if (settings.X > -30000 && settings.Y > -30000) Location = new Point(settings.X, settings.Y);
        TopMost = settings.TopMost;
        topMostButton.Checked = settings.TopMost;
    }

    private void CaptureWindowSettings()
    {
        if (WindowState == FormWindowState.Normal)
        {
            settings.X = Location.X;
            settings.Y = Location.Y;
            settings.Width = Width;
            settings.Height = Height;
        }
        settings.TopMost = TopMost;
        settings.TrackMode = trackModeBox.SelectedItem?.ToString() ?? "全部统计";
        settings.Range = rangeBox.SelectedItem?.ToString() ?? "今天";
        settings.Theme = themeBox.SelectedItem?.ToString() ?? "浅色";
        settings.MinimizeToBackground = minimizeBackgroundButton.Checked;
        settings.AutoStart = autoStartButton.Checked;
        settings.TrayIconVisible = trayButton.Checked;
        settings.OverlayVisible = overlayButton.Checked;
        settings.SelectedTabIndex = tabs.SelectedIndex;
        settings.ListZoom = listZoom;
        settings.RankSortColumn = rankSortColumn;
        settings.RankSortOrder = rankSortOrder;
        settings.GroupSortColumn = groupSortColumn;
        settings.GroupSortOrder = groupSortOrder;
        settings.AppSortColumn = appSortColumn;
        settings.AppSortOrder = appSortOrder;
        settings.OverviewTopRatio = combinedView.OverviewTopRatio;
        settings.OverviewMouseRatio = combinedView.OverviewMouseRatio;
    }

    private void ApplyTheme(string theme)
    {
        HeatmapRenderer.ApplyTheme(theme);
        var bg = HeatmapRenderer.SurfaceColor;
        var card = HeatmapRenderer.CardColor;
        var text = HeatmapRenderer.TextColor;

        BackColor = bg;
        foreach (TabPage page in tabs.TabPages) page.BackColor = bg;
        combinedView.BackColor = bg;
        keyboardView.BackColor = bg;
        mouseView.BackColor = bg;
        peakView.BackColor = bg;
        speedView.BackColor = bg;
        tabs.BackColor = bg;
        rankList.BackColor = card;
        rankList.ForeColor = text;
        statusStrip.BackColor = card;
        statusStrip.ForeColor = text;
        foreach (var strip in Controls.OfType<TableLayoutPanel>().SelectMany(p => p.Controls.OfType<ToolStrip>()))
        {
            strip.BackColor = card;
            strip.ForeColor = text;
        }
        foreach (var list in tabs.TabPages.Cast<TabPage>().SelectMany(t => t.Controls.OfType<ListView>()))
        {
            list.BackColor = card;
            list.ForeColor = text;
        }
        RefreshAll();
        Invalidate(true);
    }

    private string DataFile(DateOnly date) => Path.Combine(appDir, $"usage-{date:yyyy-MM-dd}.json");

    private UsageState LoadState(DateOnly date)
    {
        var path = DataFile(date);
        if (File.Exists(path))
        {
            try
            {
                var loaded = JsonSerializer.Deserialize<UsageState>(File.ReadAllText(path, Encoding.UTF8));
                if (loaded != null)
                {
                    loaded.Keys ??= new Dictionary<string, int>();
                    loaded.Mouse ??= new Dictionary<string, int>();
                    loaded.Combos ??= new Dictionary<string, int>();
                    loaded.RapidKeys ??= new Dictionary<string, int>();
                    loaded.AppKeys ??= new Dictionary<string, Dictionary<string, int>>();
                    loaded.AppUsageSeconds ??= new Dictionary<string, double>();
                    loaded.AppPaths ??= new Dictionary<string, string>();
                    loaded.HourlyKeys = NormalizeHourly(loaded.HourlyKeys);
                    loaded.HourlyMouse = NormalizeHourly(loaded.HourlyMouse);
                    loaded.MinuteKeys = NormalizeMinute(loaded.MinuteKeys);
                    loaded.Date = date.ToString("yyyy-MM-dd");
                    return loaded;
                }
            }
            catch
            {
                MessageBox.Show($"数据文件读取失败，将使用空数据：\n{path}", "数据错误", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
        return NewState(date);
    }

    private void RecalculateTotals()
    {
        keyTotalCache = state.Keys.Values.Sum();
        mouseTotalCache = state.Mouse.Values.Sum();
    }

    private void EnsureHourlyArrays()
    {
        state.HourlyKeys = NormalizeHourly(state.HourlyKeys);
        state.HourlyMouse = NormalizeHourly(state.HourlyMouse);
    }

    private void EnsureMinuteArrays()
    {
        state.MinuteKeys = NormalizeMinute(state.MinuteKeys);
    }

    private static int[] NormalizeHourly(int[]? values)
    {
        var normalized = new int[24];
        if (values == null) return normalized;
        Array.Copy(values, normalized, Math.Min(24, values.Length));
        return normalized;
    }

    private static int[] NormalizeMinute(int[]? values)
    {
        var normalized = new int[1440];
        if (values == null) return normalized;
        Array.Copy(values, normalized, Math.Min(1440, values.Length));
        return normalized;
    }

    private void SaveState(bool force = false)
    {
        if (!force && !stateDirty) return;
        if (!force && saveInProgress)
        {
            stateDirty = true;
            return;
        }

        Directory.CreateDirectory(appDir);
        var path = CurrentDataFile;
        var snapshot = CloneState(state);
        stateDirty = false;

        if (force)
        {
            WriteStateFile(path, snapshot);
            return;
        }

        saveInProgress = true;
        Task.Run(() => WriteStateFile(path, snapshot))
            .ContinueWith(_ =>
            {
                saveInProgress = false;
                if (stateDirty && !IsDisposed)
                {
                    try { BeginInvoke(new Action(() => SaveState())); }
                    catch { }
                }
            }, TaskScheduler.Default);
    }

    private void WriteStateFile(string path, UsageState snapshot)
    {
        lock (saveLock)
        {
            var temp = path + ".tmp";
            File.WriteAllText(temp, JsonSerializer.Serialize(snapshot, jsonOptions), Encoding.UTF8);
            File.Copy(temp, path, true);
            File.Delete(temp);
        }
    }

    private static UsageState CloneState(UsageState source) => new()
    {
        Date = source.Date,
        Keys = new Dictionary<string, int>(source.Keys ?? new Dictionary<string, int>(), StringComparer.OrdinalIgnoreCase),
        Mouse = new Dictionary<string, int>(source.Mouse ?? new Dictionary<string, int>(), StringComparer.OrdinalIgnoreCase),
        HourlyKeys = NormalizeHourly(source.HourlyKeys),
        HourlyMouse = NormalizeHourly(source.HourlyMouse),
        MinuteKeys = NormalizeMinute(source.MinuteKeys),
        Combos = new Dictionary<string, int>(source.Combos ?? new Dictionary<string, int>(), StringComparer.OrdinalIgnoreCase),
        RapidKeys = new Dictionary<string, int>(source.RapidKeys ?? new Dictionary<string, int>(), StringComparer.OrdinalIgnoreCase),
        AppKeys = (source.AppKeys ?? new Dictionary<string, Dictionary<string, int>>())
            .ToDictionary(x => x.Key, x => new Dictionary<string, int>(x.Value, StringComparer.OrdinalIgnoreCase), StringComparer.OrdinalIgnoreCase),
        AppUsageSeconds = new Dictionary<string, double>(source.AppUsageSeconds ?? new Dictionary<string, double>(), StringComparer.OrdinalIgnoreCase),
        AppPaths = new Dictionary<string, string>(source.AppPaths ?? new Dictionary<string, string>(), StringComparer.OrdinalIgnoreCase),
        MouseDistancePixels = source.MouseDistancePixels
    };

    private void MigrateOldData()
    {
        var oldFile = Path.Combine(appDir, "daily-counts.json");
        if (!File.Exists(oldFile)) return;
        var todayFile = DataFile(DateOnly.FromDateTime(DateTime.Today));
        if (!File.Exists(todayFile)) File.Copy(oldFile, todayFile);
    }

    private void MigrateVersionData()
    {
        try
        {
            var oldDir = Path.Combine(AppContext.BaseDirectory, "data");
            if (!Directory.Exists(oldDir)) return;
            if (string.Equals(Path.GetFullPath(oldDir).TrimEnd(Path.DirectorySeparatorChar), Path.GetFullPath(appDir).TrimEnd(Path.DirectorySeparatorChar), StringComparison.OrdinalIgnoreCase)) return;

            Directory.CreateDirectory(appDir);
            foreach (var file in Directory.GetFiles(oldDir, "*.json"))
            {
                var dest = Path.Combine(appDir, Path.GetFileName(file));
                if (!File.Exists(dest)) File.Copy(file, dest);
            }
        }
        catch { }
    }

    private static bool IsStartupEnabled()
    {
        using var key = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Run", false);
        return string.Equals(key?.GetValue("KeyMouseHeatmap") as string, Application.ExecutablePath, StringComparison.OrdinalIgnoreCase);
    }

    private static void SetStartupEnabled(bool enabled)
    {
        using var key = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Run", true);
        if (key == null) return;
        if (enabled) key.SetValue("KeyMouseHeatmap", Application.ExecutablePath);
        else key.DeleteValue("KeyMouseHeatmap", false);
    }

    private static UsageState NewState(DateOnly date) => new()
    {
        Date = date.ToString("yyyy-MM-dd"),
        Keys = new Dictionary<string, int>(),
        Mouse = new Dictionary<string, int>(),
        HourlyKeys = new int[24],
        HourlyMouse = new int[24],
        MinuteKeys = new int[1440],
        Combos = new Dictionary<string, int>(),
        RapidKeys = new Dictionary<string, int>(),
        AppKeys = new Dictionary<string, Dictionary<string, int>>(),
        AppUsageSeconds = new Dictionary<string, double>(),
        AppPaths = new Dictionary<string, string>(),
        MouseDistancePixels = 0
    };

    private static void AddCount(Dictionary<string, int> table, string name) => table[name] = table.TryGetValue(name, out var count) ? count + 1 : 1;
    internal static string FormatDistance(double pixels)
    {
        const double metersPerPixelAt96Dpi = 0.0254 / 96.0;
        var meters = Math.Max(0, pixels) * metersPerPixelAt96Dpi;
        if (meters < 1000) return $"{meters:F1} m";
        return $"{meters / 1000.0:F2} km";
    }

    internal static string FormatDuration(double seconds)
    {
        var totalSeconds = (long)Math.Round(Math.Max(0, seconds));
        var duration = TimeSpan.FromSeconds(totalSeconds);
        if (duration.TotalHours >= 24) return $"{(int)duration.TotalDays}天{duration.Hours:D2}小时";
        if (duration.TotalHours >= 1) return $"{(int)duration.TotalHours}小时{duration.Minutes:D2}分";
        if (duration.TotalMinutes >= 1) return $"{duration.Minutes}分{duration.Seconds:D2}秒";
        return $"{duration.Seconds}秒";
    }

    internal static string FormatDurationForUi(double seconds)
    {
        var totalSeconds = (long)Math.Floor(Math.Max(0, seconds));
        var duration = TimeSpan.FromSeconds(totalSeconds);
        if (duration.TotalHours >= 24) return $"{(int)duration.TotalDays}天{duration.Hours:D2}小时{duration.Minutes:D2}分{duration.Seconds:D2}秒";
        if (duration.TotalHours >= 1) return $"{(int)duration.TotalHours}小时{duration.Minutes:D2}分{duration.Seconds:D2}秒";
        if (duration.TotalMinutes >= 1) return $"{duration.Minutes}分{duration.Seconds:D2}秒";
        return $"{totalSeconds}秒";
    }

    private static void EnableDoubleBuffer(Control control)
    {
        try
        {
            typeof(Control).GetProperty("DoubleBuffered", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)?.SetValue(control, true);
        }
        catch
        {
            // 某些运行环境不允许反射设置时忽略，功能不受影响。
        }
    }

    private static string EscapeCsv(string value) => value.Replace("\"", "\"\"");

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool PostMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);
}


internal sealed class HeatmapTabControl : TabControl
{
    private static bool IsArrowKey(Keys keyData)
    {
        var key = keyData & Keys.KeyCode;
        return key is Keys.Left or Keys.Right or Keys.Up or Keys.Down;
    }

    protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
    {
        // WinForms 的 TabControl 默认会用方向键切换标签页。
        // 本软件方向键本身是统计对象，所以这里吃掉 TabControl 的导航行为，
        // Raw Input 仍然会正常收到按下/松开事件并实时高亮。
        if (IsArrowKey(keyData)) return true;
        return base.ProcessCmdKey(ref msg, keyData);
    }

    protected override bool IsInputKey(Keys keyData)
    {
        if (IsArrowKey(keyData)) return true;
        return base.IsInputKey(keyData);
    }

    protected override void OnKeyDown(KeyEventArgs e)
    {
        if (IsArrowKey(e.KeyData))
        {
            e.Handled = true;
            e.SuppressKeyPress = true;
            return;
        }
        base.OnKeyDown(e);
    }

    protected override void WndProc(ref Message m)
    {
        const int WM_KEYDOWN = 0x0100;
        const int WM_SYSKEYDOWN = 0x0104;
        if (m.Msg is WM_KEYDOWN or WM_SYSKEYDOWN)
        {
            var key = (Keys)((int)m.WParam) & Keys.KeyCode;
            if (key is Keys.Left or Keys.Right or Keys.Up or Keys.Down) return;
        }
        base.WndProc(ref m);
    }
}

internal sealed class CombinedView : ScrollableControl
{
    // 总览页使用响应式布局：上方显示键盘/鼠标热力图，下方显示排行、分组、应用详情 Top 10。
    private const int Gap = 12;
    private const int PaddingSize = 10;

    private readonly record struct OverviewLayout(Rectangle Keyboard, Rectangle Mouse, Rectangle Rank, Rectangle Group, Rectangle App);
    private readonly record struct OverviewRow(string Left, string Right, string? Detail, int Value, int Total, string? AuxText = null, string? IconKey = null, string? IconPath = null);

    [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
    public Func<UsageState>? StateProvider { get; set; }
    [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
    public Func<IEnumerable<string>>? PressedKeysProvider { get; set; }
    [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
    public Func<IEnumerable<string>>? PressedMouseProvider { get; set; }
    [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
    public bool IsResizing { get; set; }
    [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
    public int AppSortColumn
    {
        get => appSortColumn;
        set
        {
            if (appSortColumn == value) return;
            appSortColumn = value;
            InvalidateOverviewRowsCache();
        }
    }
    [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
    public SortOrder AppSortOrder
    {
        get => appSortOrder;
        set
        {
            if (appSortOrder == value) return;
            appSortOrder = value;
            InvalidateOverviewRowsCache();
        }
    }
    [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
    public float OverviewTopRatio
    {
        get => overviewTopRatio;
        set => overviewTopRatio = Math.Clamp(value <= 0 ? 0.68f : value, 0.48f, 0.82f);
    }
    [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
    public float OverviewMouseRatio
    {
        get => overviewMouseRatio;
        set => overviewMouseRatio = Math.Clamp(value <= 0 ? 0.23f : value, 0.16f, 0.42f);
    }
    public event EventHandler? LayoutSettingsChanged;

    private List<OverviewRow> cachedRankRows = new();
    private List<OverviewRow> cachedGroupRows = new();
    private List<OverviewRow> cachedAppRows = new();
    private readonly Dictionary<string, Bitmap> overviewIconCache = new(StringComparer.OrdinalIgnoreCase);
    private readonly Bitmap overviewDefaultAppIcon = SystemIcons.Application.ToBitmap();
    private UsageState? cachedOverviewState;
    private DateTime cachedOverviewRowsAtUtc = DateTime.MinValue;
    private bool hasOverviewRowsCache;
    private float overviewTopRatio = 0.68f;
    private float overviewMouseRatio = 0.23f;
    private int appSortColumn = 1;
    private SortOrder appSortOrder = SortOrder.Descending;
    private readonly ContextMenuStrip overviewLayoutMenu = new();
    private static readonly TimeSpan OverviewRowsRefreshInterval = TimeSpan.FromMilliseconds(550);

    public CombinedView()
    {
        DoubleBuffered = true;
        BackColor = Color.FromArgb(247, 248, 250);
        AutoScroll = true;
        AutoScrollMinSize = Size.Empty;
        BuildOverviewLayoutMenu();
        ContextMenuStrip = overviewLayoutMenu;
        MouseWheel += OverviewMouseWheel;
    }

    private void BuildOverviewLayoutMenu()
    {
        overviewLayoutMenu.Items.Add("热力图区更大 / Top10 更小", null, (_, _) => AdjustOverviewLayout(0.05f, 0f));
        overviewLayoutMenu.Items.Add("Top10 更大 / 热力图区更小", null, (_, _) => AdjustOverviewLayout(-0.05f, 0f));
        overviewLayoutMenu.Items.Add(new ToolStripSeparator());
        overviewLayoutMenu.Items.Add("键盘区更宽 / 鼠标区更窄", null, (_, _) => AdjustOverviewLayout(0f, -0.03f));
        overviewLayoutMenu.Items.Add("鼠标区更宽 / 键盘区更窄", null, (_, _) => AdjustOverviewLayout(0f, 0.03f));
        overviewLayoutMenu.Items.Add(new ToolStripSeparator());
        overviewLayoutMenu.Items.Add("重置总览版块大小", null, (_, _) => ResetOverviewLayout());
        overviewLayoutMenu.Items.Add(new ToolStripSeparator());
        overviewLayoutMenu.Items.Add("提示：Ctrl+滚轮调上下比例，Ctrl+Shift+滚轮调键盘/鼠标宽度").Enabled = false;
    }

    private void OverviewMouseWheel(object? sender, MouseEventArgs e)
    {
        if ((ModifierKeys & Keys.Control) == 0) return;
        if (e is HandledMouseEventArgs handled) handled.Handled = true;
        var step = e.Delta > 0 ? 0.03f : -0.03f;
        if ((ModifierKeys & Keys.Shift) != 0)
            AdjustOverviewLayout(0f, -step);
        else
            AdjustOverviewLayout(step, 0f);
    }

    private void AdjustOverviewLayout(float topDelta, float mouseDelta)
    {
        OverviewTopRatio = overviewTopRatio + topDelta;
        OverviewMouseRatio = overviewMouseRatio + mouseDelta;
        UpdateScrollSizeForLayout();
        Invalidate();
        LayoutSettingsChanged?.Invoke(this, EventArgs.Empty);
    }

    private void ResetOverviewLayout()
    {
        overviewTopRatio = 0.68f;
        overviewMouseRatio = 0.23f;
        UpdateScrollSizeForLayout();
        Invalidate();
        LayoutSettingsChanged?.Invoke(this, EventArgs.Empty);
    }

    private void InvalidateOverviewRowsCache()
    {
        hasOverviewRowsCache = false;
        cachedOverviewState = null;
        cachedOverviewRowsAtUtc = DateTime.MinValue;
        Invalidate();
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        base.OnPaint(e);
        if (IsResizing) return;
        e.Graphics.TranslateTransform(AutoScrollPosition.X, AutoScrollPosition.Y);

        var state = StateProvider?.Invoke() ?? new UsageState();
        var pressedKeys = PressedKeysProvider?.Invoke() ?? Array.Empty<string>();
        var pressedMouse = PressedMouseProvider?.Invoke() ?? Array.Empty<string>();
        var layout = GetLayout();

        HeatmapRenderer.DrawKeyboard(e.Graphics, layout.Keyboard, state, true, pressedKeys);
        HeatmapRenderer.DrawMouse(e.Graphics, layout.Mouse, state, pressedMouse);

        EnsureOverviewRowsCache(state);
        DrawTopListCard(e.Graphics, layout.Rank, "排行 Top 10", cachedRankRows);
        DrawTopListCard(e.Graphics, layout.Group, "分组 Top 10", cachedGroupRows);
        DrawTopListCard(e.Graphics, layout.App, "应用详情 Top 10", cachedAppRows);
    }

    private void EnsureOverviewRowsCache(UsageState state)
    {
        var now = DateTime.UtcNow;
        if (hasOverviewRowsCache && ReferenceEquals(cachedOverviewState, state) && now - cachedOverviewRowsAtUtc < OverviewRowsRefreshInterval)
            return;

        cachedOverviewState = state;
        cachedOverviewRowsAtUtc = now;
        hasOverviewRowsCache = true;
        cachedRankRows = BuildOverviewRankRows(state);
        cachedGroupRows = BuildOverviewGroupRows(state);
        cachedAppRows = BuildOverviewAppRows(state);
        EnsureOverviewIcons(cachedAppRows);
    }

    private OverviewLayout GetLayout()
    {
        var availableW = Math.Max(360, ClientSize.Width - PaddingSize * 2);
        var availableH = Math.Max(320, ClientSize.Height - PaddingSize * 2);

        if (availableW >= 930)
        {
            // 总览默认把上方热力图区域放大，下方 Top10 卡片变矮；也允许用户通过右键菜单/Ctrl+滚轮自定义比例。
            var minBottomH = Math.Clamp((int)Math.Round(availableH * 0.18), 130, 230);
            var maxTopH = Math.Max(260, availableH - minBottomH - Gap - 4);
            var minTopH = Math.Min(maxTopH, Math.Clamp((int)Math.Round(availableH * 0.42), 300, 430));
            var wantedTopH = (int)Math.Round(availableH * overviewTopRatio);
            var topH = Math.Clamp(wantedTopH, minTopH, maxTopH);
            var bottomY = PaddingSize + 4 + topH + Gap;
            var bottomH = Math.Max(minBottomH, availableH - topH - Gap - 4);

            var mouseW = Math.Clamp((int)(availableW * overviewMouseRatio), 300, Math.Min(620, availableW - 560));
            var keyboardW = availableW - mouseW - Gap;

            var topY = PaddingSize + 4;
            var keyboard = new Rectangle(PaddingSize, topY, keyboardW, topH);
            var mouse = new Rectangle(PaddingSize + keyboardW + Gap, topY, mouseW, topH);

            var cardW = (availableW - Gap * 2) / 3;
            var rank = new Rectangle(PaddingSize, bottomY, cardW, bottomH);
            var group = new Rectangle(PaddingSize + cardW + Gap, bottomY, cardW, bottomH);
            var app = new Rectangle(PaddingSize + (cardW + Gap) * 2, bottomY, availableW - (cardW + Gap) * 2, bottomH);
            return new OverviewLayout(keyboard, mouse, rank, group, app);
        }

        // 小窗口：热力图和三个 Top 10 卡片纵向排列，靠滚动保证内容完整。
        var keyboardH = Math.Clamp((int)Math.Round(availableW * 0.50), 250, 380);
        var mouseH = Math.Clamp((int)Math.Round(availableW * 0.62), 300, 480);
        var cardH = Math.Clamp((int)Math.Round(availableW * 0.34), 190, 260);
        var y = PaddingSize + 4;
        var keyboardRect = new Rectangle(PaddingSize, y, availableW, keyboardH);
        y += keyboardH + Gap;
        var mouseRect = new Rectangle(PaddingSize, y, availableW, mouseH);
        y += mouseH + Gap;
        var rankRect = new Rectangle(PaddingSize, y, availableW, cardH);
        y += cardH + Gap;
        var groupRect = new Rectangle(PaddingSize, y, availableW, cardH);
        y += cardH + Gap;
        var appRect = new Rectangle(PaddingSize, y, availableW, cardH);
        return new OverviewLayout(keyboardRect, mouseRect, rankRect, groupRect, appRect);
    }

    protected override void OnResize(EventArgs e)
    {
        UpdateScrollSizeForLayout();
        base.OnResize(e);
    }

    private void UpdateScrollSizeForLayout()
    {
        var layout = GetLayout();
        var maxRight = Math.Max(Math.Max(layout.Keyboard.Right, layout.Mouse.Right), Math.Max(layout.Rank.Right, Math.Max(layout.Group.Right, layout.App.Right)));
        var maxBottom = Math.Max(Math.Max(layout.Keyboard.Bottom, layout.Mouse.Bottom), Math.Max(layout.Rank.Bottom, Math.Max(layout.Group.Bottom, layout.App.Bottom)));
        AutoScrollMinSize = new Size(
            Math.Max(0, maxRight + PaddingSize),
            Math.Max(0, maxBottom + PaddingSize));
    }

    private void DrawTopListCard(Graphics g, Rectangle bounds, string title, IReadOnlyList<OverviewRow> rows)
    {
        if (bounds.Width < 80 || bounds.Height < 90) return;

        g.SmoothingMode = SmoothingMode.AntiAlias;
        using var bg = new SolidBrush(HeatmapRenderer.CardColor);
        using var border = new Pen(HeatmapRenderer.BorderColor);
        using var path = RoundedPath(bounds, 8);
        g.FillPath(bg, path);
        g.DrawPath(border, path);

        var isAppCard = title.Contains("应用详情", StringComparison.Ordinal);
        var pad = Math.Clamp(bounds.Width / 42, 8, 16);
        var titleH = Math.Clamp((int)Math.Round(bounds.Height * 0.16), 28, 44);
        using var titleFont = new Font("Microsoft YaHei UI", Math.Clamp(bounds.Height * 0.055F, 10F, 17F), FontStyle.Bold);
        DrawFittedText(g, title, titleFont, HeatmapRenderer.TextColor,
            new Rectangle(bounds.X + pad, bounds.Y + 8, bounds.Width - pad * 2, titleH),
            TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis | TextFormatFlags.SingleLine);

        var listTop = bounds.Y + 8 + titleH + 2;
        var listBottom = bounds.Bottom - pad;
        var usableH = Math.Max(20, listBottom - listTop);
        var preferredRowH = bounds.Width >= 360 ? 52 : 46;
        var maxRows = Math.Min(10, Math.Max(1, usableH / preferredRowH));
        var visibleRows = rows.Count == 0 ? 1 : Math.Min(rows.Count, maxRows);
        var rowH = Math.Max(34, usableH / Math.Max(1, visibleRows));

        if (rows.Count == 0)
        {
            using var emptyFont = new Font("Microsoft YaHei UI", Math.Clamp(rowH * 0.40F, 8.5F, 13F), FontStyle.Regular);
            DrawFittedText(g, "暂无数据", emptyFont, HeatmapRenderer.MutedColor,
                new Rectangle(bounds.X + pad, listTop, bounds.Width - pad * 2, rowH),
                TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis | TextFormatFlags.SingleLine);
            return;
        }

        using var linePen = new Pen(Color.FromArgb(70, HeatmapRenderer.BorderColor));
        for (var i = 0; i < visibleRows; i++)
        {
            var row = rows[i];
            var rowRect = new Rectangle(bounds.X + pad, listTop + i * rowH, bounds.Width - pad * 2, rowH);
            if (i > 0) g.DrawLine(linePen, rowRect.Left, rowRect.Top, rowRect.Right, rowRect.Top);

            var rightW = Math.Clamp(bounds.Width / 3, 68, 154);
            var hasDetail = !string.IsNullOrWhiteSpace(row.Detail);
            var barH = Math.Clamp(rowH / 5, 7, 12);
            var detailH = hasDetail ? Math.Clamp(rowH / 3, 13, 22) : 0;
            var mainH = Math.Max(16, rowH - detailH - barH - 5);
            var iconSize = row.IconKey != null ? Math.Clamp(mainH - 2, 16, 22) : 0;
            var leftPad = iconSize > 0 ? iconSize + 8 : 0;
            var mainRect = new Rectangle(rowRect.X, rowRect.Y + 1, rowRect.Width, mainH);

            if (iconSize > 0)
            {
                var iconRect = new Rectangle(mainRect.X, mainRect.Y + Math.Max(0, (mainRect.Height - iconSize) / 2), iconSize, iconSize);
                var icon = GetOverviewIcon(row);
                if (icon != null)
                    g.DrawImage(icon, iconRect);
            }

            using var mainFont = new Font("Microsoft YaHei UI", Math.Clamp(mainH * 0.44F, 8F, 13F), FontStyle.Bold);
            using var rightFont = new Font("Segoe UI", Math.Clamp(mainH * 0.40F, 8F, 13F), FontStyle.Bold);
            using var auxFont = new Font("Microsoft YaHei UI", Math.Clamp(mainH * 0.34F, 7F, 11F), FontStyle.Regular);

            var rightRect = new Rectangle(rowRect.Right - rightW, mainRect.Y, rightW, mainRect.Height);
            if (isAppCard)
            {
                var auxW = Math.Clamp(bounds.Width / 5, 70, 130);
                var auxRect = new Rectangle(rightRect.X - auxW - 6, mainRect.Y, auxW, mainRect.Height);
                var leftRect = new Rectangle(mainRect.X + leftPad, mainRect.Y, Math.Max(40, auxRect.X - (mainRect.X + leftPad) - 6), mainRect.Height);
                DrawFittedText(g, row.Left, mainFont, HeatmapRenderer.TextColor,
                    leftRect,
                    TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis | TextFormatFlags.SingleLine);
                if (!string.IsNullOrWhiteSpace(row.AuxText))
                {
                    DrawFittedText(g, row.AuxText!, auxFont, HeatmapRenderer.MutedColor,
                        auxRect,
                        TextFormatFlags.Right | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis | TextFormatFlags.SingleLine);
                }
            }
            else
            {
                DrawFittedText(g, row.Left, mainFont, HeatmapRenderer.TextColor,
                    new Rectangle(mainRect.X + leftPad, mainRect.Y, mainRect.Width - rightW - 8 - leftPad, mainRect.Height),
                    TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis | TextFormatFlags.SingleLine);
            }
            DrawFittedText(g, row.Right, rightFont, HeatmapRenderer.HotColor,
                rightRect,
                TextFormatFlags.Right | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis | TextFormatFlags.SingleLine);

            var barRect = row.Total > 0
                ? new Rectangle(rowRect.X + leftPad, rowRect.Bottom - barH - 2, rowRect.Width - leftPad, barH)
                : Rectangle.Empty;

            if (hasDetail)
            {
                var detailTop = isAppCard && !barRect.IsEmpty
                    ? Math.Max(mainRect.Bottom, barRect.Y - detailH - 2)
                    : mainRect.Bottom;
                using var detailFont = new Font("Microsoft YaHei UI", Math.Clamp(detailH * 0.66F, 8.5F, 12.5F), FontStyle.Regular);
                var detailFlags = isAppCard
                    ? TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis | TextFormatFlags.SingleLine
                    : TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis | TextFormatFlags.SingleLine;
                var detailRect = isAppCard && !barRect.IsEmpty
                    ? new Rectangle(barRect.X, detailTop, barRect.Width, detailH)
                    : new Rectangle(rowRect.X + leftPad, detailTop, rowRect.Width - leftPad, detailH);
                DrawFittedText(g, row.Detail!, detailFont, HeatmapRenderer.MutedColor, detailRect, detailFlags);
            }

            if (!barRect.IsEmpty)
                DrawOverviewStatBar(g, barRect, row.Value, row.Total);
        }
    }

    private static List<OverviewRow> BuildOverviewRankRows(UsageState state)
    {
        var total = state.Keys.Values.Sum() + state.Mouse.Values.Sum();
        return state.Keys.Select(x => (Device: "键盘", Name: x.Key, Count: x.Value))
            .Concat(state.Mouse.Select(x => (Device: "鼠标", Name: x.Key, Count: x.Value)))
            .Where(x => x.Count > 0)
            .OrderByDescending(x => x.Count)
            .ThenBy(x => x.Name)
            .Take(10)
            .Select((x, i) => new OverviewRow($"{i + 1}. {x.Name}", $"{x.Count}", x.Device, x.Count, total))
            .ToList();
    }

    private static List<OverviewRow> BuildOverviewGroupRows(UsageState s)
    {
        int Sum(params string[] names) => names.Sum(n => s.Keys.TryGetValue(n, out var v) ? v : 0);
        var letters = s.Keys.Where(kv => kv.Key.Length == 1 && kv.Key[0] >= 'A' && kv.Key[0] <= 'Z').Sum(kv => kv.Value);
        var nums = s.Keys.Where(kv => kv.Key.Length == 1 && kv.Key[0] >= '0' && kv.Key[0] <= '9').Sum(kv => kv.Value);
        var fn = s.Keys.Where(kv => kv.Key.StartsWith("F") && kv.Key.Length > 1 && int.TryParse(kv.Key[1..], out _)).Sum(kv => kv.Value);
        var numpad = s.Keys.Where(kv => kv.Key.StartsWith("Num ")).Sum(kv => kv.Value);

        var rows = new[]
        {
            ("字母区", letters, "A-Z 主键区字母"),
            ("主键数字区", nums, "键盘上方 0-9"),
            ("功能键区", fn, "F1-F12"),
            ("方向键区", Sum("Up", "Down", "Left", "Right"), "上、下、左、右"),
            ("小数字键盘", numpad, "Num 0-9 / 运算符"),
            ("修饰键", Sum("Left Shift", "Right Shift", "Left Ctrl", "Right Ctrl", "Left Alt", "Right Alt", "Left Win", "Right Win", "Menu"), "Shift/Ctrl/Alt/Win"),
            ("鼠标", s.Mouse.Values.Sum(), "左键、右键、中键、侧键、滚轮"),
            ("组合键", s.Combos?.Values.Sum() ?? 0, "同时按下两个及以上键"),
            ("连击", s.RapidKeys?.Values.Sum() ?? 0, "1 秒内重复点击同一个键")
        };

        var total = rows.Sum(x => x.Item2);
        return rows
            .Where(x => x.Item2 > 0)
            .OrderByDescending(x => x.Item2)
            .ThenBy(x => x.Item1)
            .Take(10)
            .Select((x, i) => new OverviewRow($"{i + 1}. {x.Item1}", $"{x.Item2}", x.Item3, x.Item2, total))
            .ToList();
    }

    private List<OverviewRow> BuildOverviewAppRows(UsageState s)
    {
        var appKeys = s.AppKeys ?? new Dictionary<string, Dictionary<string, int>>(StringComparer.OrdinalIgnoreCase);
        var appUsage = s.AppUsageSeconds ?? new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);
        var names = appKeys.Keys
            .Concat(appUsage.Keys)
            .Where(x => !string.IsNullOrWhiteSpace(x))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .GroupBy(NormalizeOverviewAppName, StringComparer.OrdinalIgnoreCase);

        var mergedApps = names
            .Select(group =>
            {
                var totalCount = 0;
                var keyTypes = 0;
                var usage = 0.0;
                foreach (var name in group)
                {
                    if (appKeys.TryGetValue(name, out var keys))
                    {
                        totalCount += keys.Values.Sum();
                        keyTypes += keys.Count(x => x.Value > 0);
                    }
                    if (appUsage.TryGetValue(name, out var seconds))
                        usage += seconds;
                }

                var iconPath = group
                    .Select(name => TryGetOverviewAppPath(s, name))
                    .FirstOrDefault(path => !string.IsNullOrWhiteSpace(path));
                iconPath ??= AppIconProvider.FindRunningProcessPath(group.Key);

                var summary = totalCount > 0
                    ? string.Join(", ", group
                        .SelectMany(name => appKeys.TryGetValue(name, out var keys) ? keys : new Dictionary<string, int>())
                        .GroupBy(x => x.Key, StringComparer.OrdinalIgnoreCase)
                        .Select(g => new KeyValuePair<string,int>(g.Key, g.Sum(v => v.Value)))
                        .Where(x => x.Value > 0)
                        .OrderByDescending(x => x.Value)
                        .ThenBy(x => x.Key)
                        .Take(5)
                        .Select(x => $"{x.Key} {x.Value}"))
                    : string.Empty;

                var appName = group.Key;
                return new
                {
                    App = appName,
                    Count = totalCount,
                    KeyTypes = keyTypes,
                    Usage = usage,
                    IconPath = iconPath,
                    IconKey = AppIconProvider.IconKey(appName, iconPath),
                    Summary = summary
                };
            })
            .Where(x => x.Count > 0 || x.Usage >= 1)
            .ToList();

        var orderedTop = (appSortColumn switch
        {
            0 => appSortOrder == SortOrder.Ascending ? mergedApps.OrderBy(x => x.App) : mergedApps.OrderByDescending(x => x.App),
            1 => appSortOrder == SortOrder.Ascending ? mergedApps.OrderBy(x => x.Usage).ThenBy(x => x.App) : mergedApps.OrderByDescending(x => x.Usage).ThenByDescending(x => x.Count).ThenBy(x => x.App),
            2 or 4 => appSortOrder == SortOrder.Ascending ? mergedApps.OrderBy(x => x.Count).ThenBy(x => x.App) : mergedApps.OrderByDescending(x => x.Count).ThenBy(x => x.App),
            3 => appSortOrder == SortOrder.Ascending ? mergedApps.OrderBy(x => x.KeyTypes).ThenBy(x => x.App) : mergedApps.OrderByDescending(x => x.KeyTypes).ThenBy(x => x.App),
            5 => appSortOrder == SortOrder.Ascending ? mergedApps.OrderBy(x => x.Summary).ThenBy(x => x.App) : mergedApps.OrderByDescending(x => x.Summary).ThenBy(x => x.App),
            _ => mergedApps.OrderByDescending(x => x.Usage).ThenByDescending(x => x.Count).ThenBy(x => x.App)
        }).Take(10).ToList();
        var total = Math.Max(1, orderedTop.Sum(x => x.Count));
        return orderedTop
            .Select((x, i) => new OverviewRow(
                $"{i + 1}. {x.App}",
                $"{x.Count}次",
                BuildOverviewAppDetail(x.Count, x.KeyTypes, x.Summary),
                x.Count,
                total,
                MainForm.FormatDurationForUi(x.Usage),
                x.IconKey,
                x.IconPath))
            .ToList();
    }

    private static string BuildOverviewAppDetail(int count, int keyTypes, string summary)
    {
        if (count <= 0)
            return "暂无按键记录";
        if (string.IsNullOrWhiteSpace(summary))
            return $"{keyTypes} 种按键";
        return $"{keyTypes} 种按键 · Top5：{summary}";
    }

    private static string NormalizeOverviewAppName(string appName)
    {
        var trimmed = appName.Trim();
        var withoutExtension = Path.GetFileNameWithoutExtension(trimmed);
        return string.IsNullOrWhiteSpace(withoutExtension) ? trimmed : withoutExtension;
    }

    private static string? TryGetOverviewAppPath(UsageState source, string appName)
    {
        if (source.AppPaths != null && source.AppPaths.TryGetValue(appName, out var path) && File.Exists(path)) return path;
        var normalized = NormalizeOverviewAppName(appName);
        return source.AppPaths?
            .Where(x => string.Equals(NormalizeOverviewAppName(x.Key), normalized, StringComparison.OrdinalIgnoreCase) && File.Exists(x.Value))
            .Select(x => x.Value)
            .FirstOrDefault();
    }

    private void EnsureOverviewIcons(IEnumerable<OverviewRow> rows)
    {
        foreach (var row in rows)
        {
            if (string.IsNullOrWhiteSpace(row.IconKey) || overviewIconCache.ContainsKey(row.IconKey)) continue;
            overviewIconCache[row.IconKey] = AppIconProvider.GetIconBitmap(row.IconPath);
        }
    }

    private Bitmap GetOverviewIcon(OverviewRow row)
    {
        if (!string.IsNullOrWhiteSpace(row.IconKey) && overviewIconCache.TryGetValue(row.IconKey, out var cached))
            return cached;
        return overviewDefaultAppIcon;
    }

    private static void DrawOverviewStatBar(Graphics g, Rectangle bounds, int value, int max)
    {
        if (bounds.Width <= 12 || bounds.Height <= 4 || max <= 0) return;

        using var back = new SolidBrush(Color.FromArgb(236, 240, 245));
        using var fill = new SolidBrush(Color.FromArgb(55, 132, 245));
        using var border = new Pen(Color.FromArgb(197, 208, 222));
        g.FillRectangle(back, bounds);
        g.DrawRectangle(border, bounds);

        var ratio = Math.Clamp(value / (double)max, 0, 1);
        var fillRect = new Rectangle(bounds.X + 1, bounds.Y + 1, Math.Max(0, (int)Math.Round((bounds.Width - 2) * ratio)), Math.Max(0, bounds.Height - 2));
        g.FillRectangle(fill, fillRect);
    }

    private static GraphicsPath RoundedPath(Rectangle rect, int radius)
    {
        var path = new GraphicsPath();
        var diameter = Math.Max(2, radius * 2);
        var arc = new Rectangle(rect.X, rect.Y, diameter, diameter);
        path.AddArc(arc, 180, 90);
        arc.X = rect.Right - diameter;
        path.AddArc(arc, 270, 90);
        arc.Y = rect.Bottom - diameter;
        path.AddArc(arc, 0, 90);
        arc.X = rect.X;
        path.AddArc(arc, 90, 90);
        path.CloseFigure();
        return path;
    }

    private static void DrawFittedText(Graphics g, string value, Font font, Color color, Rectangle rect, TextFormatFlags flags)
    {
        if (rect.Width <= 0 || rect.Height <= 0 || string.IsNullOrEmpty(value)) return;

        var drawFlags = flags | TextFormatFlags.NoPadding;
        var measureFlags = drawFlags & ~TextFormatFlags.EndEllipsis & ~TextFormatFlags.WordEllipsis;
        var proposed = new Size(Math.Max(1, rect.Width), Math.Max(1, rect.Height));
        var measured = TextRenderer.MeasureText(value, font, proposed, measureFlags);
        if (measured.Width <= rect.Width && measured.Height <= rect.Height)
        {
            TextRenderer.DrawText(g, value, font, rect, color, drawFlags);
            return;
        }

        var scaleW = measured.Width <= 0 ? 1F : rect.Width / (float)measured.Width;
        var scaleH = measured.Height <= 0 ? 1F : rect.Height / (float)measured.Height;
        var scaledSize = Math.Max(5.5F, Math.Min(font.Size, font.Size * Math.Min(scaleW, scaleH) * 0.94F));
        using var fitted = new Font(font.FontFamily, scaledSize, font.Style);
        TextRenderer.DrawText(g, value, fitted, rect, color, drawFlags);
    }
}


internal sealed class KeyboardView : ScrollableControl
{
    [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
    public Func<UsageState>? StateProvider { get; set; }
    [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
    public Func<IEnumerable<string>>? PressedKeysProvider { get; set; }
    [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
    public bool IsResizing { get; set; }

    public KeyboardView()
    {
        DoubleBuffered = true;
        BackColor = Color.FromArgb(247, 248, 250);
        AutoScroll = false;
        AutoScrollMinSize = Size.Empty;
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        base.OnPaint(e);
        if (IsResizing) return;

        var state = StateProvider?.Invoke() ?? new UsageState();
        var pressed = PressedKeysProvider?.Invoke() ?? Array.Empty<string>();

        // 键盘页不再使用固定 1320px 宽度和滚动条。
        // 小窗口时整体缩小，大窗口时整体放大，始终尽量完整显示 104 键布局。
        var rect = new Rectangle(
            10,
            10,
            Math.Max(260, ClientSize.Width - 20),
            Math.Max(210, ClientSize.Height - 20));

        HeatmapRenderer.DrawKeyboard(e.Graphics, rect, state, false, pressed);
    }

    protected override void OnResize(EventArgs e)
    {
        AutoScrollMinSize = Size.Empty;
        Invalidate();
        base.OnResize(e);
    }
}

internal sealed class MouseView : ScrollableControl
{
    [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
    public Func<UsageState>? StateProvider { get; set; }
    [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
    public Func<IEnumerable<string>>? PressedKeysProvider { get; set; }
    [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
    public Func<IEnumerable<string>>? PressedMouseProvider { get; set; }
    [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
    public bool IsResizing { get; set; }

    public MouseView()
    {
        DoubleBuffered = true;
        BackColor = Color.FromArgb(247, 248, 250);
        AutoScroll = false;
        AutoScrollMinSize = Size.Empty;
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        base.OnPaint(e);
        if (IsResizing) return;
        var state = StateProvider?.Invoke() ?? new UsageState();
        var pressedMouse = PressedMouseProvider?.Invoke() ?? Array.Empty<string>();
        var rect = new Rectangle(10, 10, Math.Max(260, ClientSize.Width - 20), Math.Max(240, ClientSize.Height - 20));
        HeatmapRenderer.DrawMouse(e.Graphics, rect, state, pressedMouse);
    }

    protected override void OnResize(EventArgs e)
    {
        AutoScrollMinSize = Size.Empty;
        Invalidate();
        base.OnResize(e);
    }
}

internal sealed class PeakView : ScrollableControl
{
    [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
    public Func<UsageState>? StateProvider { get; set; }
    [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
    public Func<IEnumerable<string>>? PressedKeysProvider { get; set; }
    [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
    public bool IsResizing { get; set; }

    public PeakView()
    {
        DoubleBuffered = true;
        BackColor = Color.FromArgb(247, 248, 250);
        AutoScroll = false;
        AutoScrollMinSize = Size.Empty;
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        base.OnPaint(e);
        if (IsResizing) return;
        var state = StateProvider?.Invoke() ?? new UsageState();
        var rect = new Rectangle(10, 10, Math.Max(260, ClientSize.Width - 20), Math.Max(220, ClientSize.Height - 20));
        HeatmapRenderer.DrawPeakChart(e.Graphics, rect, state);
    }

    protected override void OnResize(EventArgs e)
    {
        AutoScrollMinSize = Size.Empty;
        Invalidate();
        base.OnResize(e);
    }
}

internal sealed class SpeedView : ScrollableControl
{
    [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
    public Func<UsageState>? StateProvider { get; set; }
    [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
    public Func<SpeedStats>? SpeedProvider { get; set; }
    [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
    public bool IsResizing { get; set; }

    public SpeedView()
    {
        DoubleBuffered = true;
        BackColor = Color.FromArgb(247, 248, 250);
        AutoScroll = false;
        AutoScrollMinSize = Size.Empty;
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        base.OnPaint(e);
        if (IsResizing) return;
        var state = StateProvider?.Invoke() ?? new UsageState();
        var liveSpeed = SpeedProvider?.Invoke() ?? default;
        var rect = new Rectangle(10, 10, Math.Max(280, ClientSize.Width - 20), Math.Max(260, ClientSize.Height - 20));
        HeatmapRenderer.DrawSpeed(e.Graphics, rect, state, liveSpeed);
    }

    protected override void OnResize(EventArgs e)
    {
        AutoScrollMinSize = Size.Empty;
        Invalidate();
        base.OnResize(e);
    }
}


internal readonly record struct SpeedStats(int CurrentMinute, int CurrentHour, int MaxMinute, int MaxHour, double ActiveAverageMinute, double AverageMinute, double AverageHour);
internal readonly record struct RankRow(int Rank, string Device, string Name, int Count);
internal readonly record struct AppKeyRow(string App, string Summary, int Count, int KeyTypes, double UsageSeconds, string IconKey, string? IconPath);
internal readonly record struct GroupBarTag(int Count, int Total);
internal readonly record struct ForegroundAppInfo(string Name, string? ExecutablePath);
internal readonly record struct InputEvent(InputEventKind Kind, string Name);

internal enum InputEventKind
{
    KeyDown,
    KeyUp,
    MouseDown,
    MouseUp
}

internal sealed class StrongToolStripRenderer : ToolStripProfessionalRenderer
{
    protected override void OnRenderButtonBackground(ToolStripItemRenderEventArgs e)
    {
        if (e.Item is not ToolStripButton button)
        {
            base.OnRenderButtonBackground(e);
            return;
        }

        var rect = new Rectangle(Point.Empty, e.Item.Size);
        var selected = button.Checked || button.Selected || button.Pressed;
        var fill = button.Checked
            ? Color.FromArgb(217, 236, 255)
            : button.Pressed
                ? Color.FromArgb(195, 223, 255)
                : button.Selected
                    ? Color.FromArgb(235, 244, 255)
                    : Color.Transparent;
        if (selected)
        {
            using var brush = new SolidBrush(fill);
            using var pen = new Pen(button.Checked ? Color.FromArgb(24, 119, 242) : Color.FromArgb(130, 180, 235), button.Checked ? 2f : 1f);
            e.Graphics.FillRectangle(brush, rect);
            e.Graphics.DrawRectangle(pen, 0, 0, rect.Width - 1, rect.Height - 1);
        }

        if (button.Checked)
        {
            using var accent = new SolidBrush(Color.FromArgb(24, 119, 242));
            e.Graphics.FillRectangle(accent, 0, rect.Height - 3, rect.Width, 3);
        }
    }
}

internal sealed class OverlayForm : Form
{
    private string keyText = "等待按键";
    private string countText = "累计 0 次";
    private string pressedText = "正在按住：无";
    private string speedText = "速度：0 次/分钟 | 0 次/小时";
    private readonly ContextMenuStrip menu = new();
    private bool showKey = true;
    private bool showSpeed = true;
    private bool showPressedKeys = true;
    private bool showBorder = true;
    private byte backgroundAlpha = 225;
    private readonly Color backgroundColor = Color.FromArgb(24, 30, 38);
    private readonly Color borderColor = Color.FromArgb(93, 211, 255);
    private readonly Color countColor = Color.FromArgb(255, 214, 142);
    private readonly Color keyColor = Color.White;
    private readonly Color pressedColor = Color.FromArgb(196, 207, 222);
    private readonly Color speedColor = Color.FromArgb(130, 235, 255);
    private readonly System.Windows.Forms.Timer renderTimer = new() { Interval = 10 };
    private bool renderPending;
    private bool isInteractiveMoveOrResize;

    public bool IsInteractiveMoveOrResize => isInteractiveMoveOrResize;

    public OverlayForm()
    {
        FormBorderStyle = FormBorderStyle.None;
        ShowInTaskbar = false;
        TopMost = true;
        BackColor = Color.Black;
        MinimumSize = new Size(190, 108);
        Size = new Size(360, 178);
        StartPosition = FormStartPosition.Manual;
        Location = new Point(Screen.PrimaryScreen?.WorkingArea.Right - 350 ?? 900, 90);
        Text = "按键悬浮窗";
        BuildMenu();
        ContextMenuStrip = menu;
        renderTimer.Tick += (_, _) => FlushRenderIfPossible();
        MouseDoubleClick += (_, _) => ToggleBorder();
        Resize += (_, _) => RenderLayeredWindow();
        Shown += (_, _) => RenderLayeredWindow();
    }

    protected override CreateParams CreateParams
    {
        get
        {
            const int WS_EX_LAYERED = 0x00080000;
            var cp = base.CreateParams;
            cp.ExStyle |= WS_EX_LAYERED;
            return cp;
        }
    }

    private void BuildMenu()
    {
        menu.Items.Add("显示边框/隐藏边框", null, (_, _) => ToggleBorder());
        menu.Items.Add(new ToolStripSeparator());
        var keyItem = new ToolStripMenuItem("悬浮窗显示跳出按键", null, (_, _) => { showKey = !showKey; ApplyOverlayVisibility(); }) { Checked = showKey, CheckOnClick = true };
        var pressedItem = new ToolStripMenuItem("悬浮窗显示正在按住", null, (_, _) => { showPressedKeys = !showPressedKeys; ApplyOverlayVisibility(); }) { Checked = showPressedKeys, CheckOnClick = true };
        var speedItem = new ToolStripMenuItem("悬浮窗显示打字速度", null, (_, _) => { showSpeed = !showSpeed; ApplyOverlayVisibility(); }) { Checked = showSpeed, CheckOnClick = true };
        menu.Items.Add(keyItem);
        menu.Items.Add(pressedItem);
        menu.Items.Add(speedItem);
        menu.Items.Add(new ToolStripSeparator());
        foreach (var item in new (string Text, byte Alpha)[]
        {
            ("背景透明度 100%", 255), ("背景透明度 90%", 230), ("背景透明度 80%", 204),
            ("背景透明度 70%", 179), ("背景透明度 60%", 153), ("背景透明度 50%", 128),
            ("背景透明度 40%", 102), ("背景透明度 30%", 77)
        })
        {
            menu.Items.Add(item.Text, null, (_, _) => { backgroundAlpha = item.Alpha; RenderLayeredWindow(); });
        }
        menu.Items.Add(new ToolStripSeparator());
        menu.Items.Add("放大一点", null, (_, _) => Size = new Size(Width + 40, Height + 24));
        menu.Items.Add("缩小一点", null, (_, _) => Size = new Size(Math.Max(MinimumSize.Width, Width - 40), Math.Max(MinimumSize.Height, Height - 24)));
    }

    private void ToggleBorder()
    {
        showBorder = !showBorder;
        RenderLayeredWindow();
    }

    public void UpdateKey(string name, int count, IEnumerable<string> pressed, SpeedStats speed, bool immediate = false)
    {
        keyText = name;
        countText = $"累计 {count} 次";
        pressedText = BuildPressedText(pressed);
        speedText = BuildSpeedText(speed);
        RenderLayeredWindow(immediate);
    }

    public void UpdateSpeed(SpeedStats speed)
    {
        var next = BuildSpeedText(speed);
        if (speedText == next) return;
        speedText = next;
        RenderLayeredWindow();
    }

    public void SetPressedKeys(IEnumerable<string> pressed, bool immediate = false)
    {
        var next = BuildPressedText(pressed);
        if (pressedText == next) return;
        pressedText = next;
        RenderLayeredWindow(immediate);
    }

    private static string BuildPressedText(IEnumerable<string> pressed)
    {
        var keys = pressed.Take(12).ToArray();
        return keys.Length == 0 ? "正在按住：无" : "正在按住：" + string.Join(" + ", keys);
    }

    private static string BuildSpeedText(SpeedStats speed) =>
        $"速度：{speed.CurrentMinute} 次/分钟 | {speed.CurrentHour} 次/小时 | 平均：{speed.AverageMinute:F1} 次/分钟";

    private void ApplyOverlayVisibility()
    {
        RenderLayeredWindow();
    }

    protected override void OnFormClosed(FormClosedEventArgs e)
    {
        renderTimer.Stop();
        renderTimer.Dispose();
        base.OnFormClosed(e);
    }

    protected override void OnMouseDown(MouseEventArgs e)
    {
        base.OnMouseDown(e);
        if (e.Button == MouseButtons.Left)
        {
            ReleaseCapture();
            SendMessage(Handle, WM_NCLBUTTONDOWN, (IntPtr)HTCAPTION, IntPtr.Zero);
        }
    }

    protected override void OnMouseUp(MouseEventArgs e)
    {
        base.OnMouseUp(e);
        if (e.Button == MouseButtons.Right)
            menu.Show(this, e.Location);
    }

    protected override void WndProc(ref Message m)
    {
        const int WM_NCHITTEST = 0x0084;
        const int WM_ENTERSIZEMOVE = 0x0231;
        const int WM_EXITSIZEMOVE = 0x0232;

        if (m.Msg == WM_ENTERSIZEMOVE)
        {
            isInteractiveMoveOrResize = true;
            base.WndProc(ref m);
            return;
        }

        if (m.Msg == WM_EXITSIZEMOVE)
        {
            base.WndProc(ref m);
            isInteractiveMoveOrResize = false;
            FlushRenderIfPossible(force: true);
            return;
        }

        if (m.Msg == WM_NCHITTEST)
        {
            base.WndProc(ref m);
            var p = PointToClient(new Point((short)(m.LParam.ToInt64() & 0xffff), (short)((m.LParam.ToInt64() >> 16) & 0xffff)));
            var grip = Math.Max(8, Math.Min(16, Math.Min(Width, Height) / 8));
            var left = p.X <= grip;
            var right = p.X >= Width - grip;
            var top = p.Y <= grip;
            var bottom = p.Y >= Height - grip;
            if (top && left) { m.Result = (IntPtr)HTTOPLEFT; return; }
            if (top && right) { m.Result = (IntPtr)HTTOPRIGHT; return; }
            if (bottom && left) { m.Result = (IntPtr)HTBOTTOMLEFT; return; }
            if (bottom && right) { m.Result = (IntPtr)HTBOTTOMRIGHT; return; }
            if (left) { m.Result = (IntPtr)HTLEFT; return; }
            if (right) { m.Result = (IntPtr)HTRIGHT; return; }
            if (top) { m.Result = (IntPtr)HTTOP; return; }
            if (bottom) { m.Result = (IntPtr)HTBOTTOM; return; }
            // 中间区域返回 HTCLIENT，避免双击被 Windows 当成标题栏双击从而最大化/全屏。
            // 拖动窗口由 OnMouseDown 主动发送 HTCAPTION 完成，右键菜单也能正常弹出。
            m.Result = (IntPtr)HTCLIENT;
            return;
        }
        base.WndProc(ref m);
    }

    private void RenderLayeredWindow(bool immediate = false)
    {
        if (IsDisposed || !IsHandleCreated || Width <= 0 || Height <= 0) return;
        renderPending = true;
        if (isInteractiveMoveOrResize) return;

        if (immediate)
        {
            FlushRenderIfPossible(force: true);
            return;
        }

        if (!renderTimer.Enabled) renderTimer.Start();
    }

    private void FlushRenderIfPossible(bool force = false)
    {
        if (!force && isInteractiveMoveOrResize) return;
        if (!renderPending)
        {
            renderTimer.Stop();
            return;
        }

        renderTimer.Stop();
        renderPending = false;
        RenderLayeredWindowNow();
    }

    private void RenderLayeredWindowNow()
    {
        if (IsDisposed || !IsHandleCreated || Width <= 0 || Height <= 0) return;
        using var bitmap = new Bitmap(Width, Height, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
        using (var g = Graphics.FromImage(bitmap))
        {
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAliasGridFit;
            g.Clear(Color.Transparent);
            var rect = new Rectangle(0, 0, Width - 1, Height - 1);
            using (var path = RoundedRect(rect, 16))
            using (var brush = new SolidBrush(Color.FromArgb(backgroundAlpha, backgroundColor)))
            {
                g.FillPath(brush, path);
                if (showBorder)
                {
                    using var pen = new Pen(Color.FromArgb(230, borderColor), 2);
                    g.DrawPath(pen, path);
                }
            }

            var rows = new List<(string Text, float Weight, Color Color, FontStyle Style)>();
            if (showKey)
            {
                rows.Add((countText, 0.85f, countColor, FontStyle.Bold));
                rows.Add((keyText, 2.0f, keyColor, FontStyle.Bold));
            }
            if (showPressedKeys) rows.Add((pressedText, 0.85f, pressedColor, FontStyle.Regular));
            if (showSpeed) rows.Add((speedText, 0.85f, speedColor, FontStyle.Regular));
            if (rows.Count == 0) rows.Add(("悬浮窗", 1.0f, keyColor, FontStyle.Bold));

            var padding = Math.Max(8, Math.Min(16, Width / 24));
            var content = new RectangleF(padding, padding, Width - padding * 2, Height - padding * 2);
            var totalWeight = rows.Sum(r => r.Weight);
            var y = content.Top;
            using var format = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center, Trimming = StringTrimming.EllipsisCharacter };
            foreach (var row in rows)
            {
                var h = content.Height * row.Weight / totalWeight;
                var rowRect = new RectangleF(content.Left, y, content.Width, h);
                var maxSize = Math.Max(8, (int)(h * 0.62f));
                var minSize = row.Text == keyText ? 12 : 6;
                using var font = GetFastFittingFont(g, row.Text, rowRect.Size, row.Style, minSize, maxSize);
                using var textBrush = new SolidBrush(row.Color);
                g.DrawString(row.Text, font, textBrush, rowRect, format);
                y += h;
            }
        }

        var screenDc = GetDC(IntPtr.Zero);
        var memDc = CreateCompatibleDC(screenDc);
        var hBitmap = bitmap.GetHbitmap(Color.FromArgb(0));
        var oldBitmap = SelectObject(memDc, hBitmap);
        try
        {
            var size = new SIZE(Width, Height);
            var pointSource = new POINT(0, 0);
            var topPos = new POINT(Left, Top);
            var blend = new BLENDFUNCTION
            {
                BlendOp = AC_SRC_OVER,
                BlendFlags = 0,
                SourceConstantAlpha = 255,
                AlphaFormat = AC_SRC_ALPHA
            };
            UpdateLayeredWindow(Handle, screenDc, ref topPos, ref size, memDc, ref pointSource, 0, ref blend, ULW_ALPHA);
        }
        finally
        {
            SelectObject(memDc, oldBitmap);
            DeleteObject(hBitmap);
            DeleteDC(memDc);
            ReleaseDC(IntPtr.Zero, screenDc);
        }
    }

    private static Font GetFastFittingFont(Graphics g, string text, SizeF box, FontStyle style, int minSize, int maxSize)
    {
        var family = text == "等待按键" || text.Length <= 12 ? "Segoe UI" : "Microsoft YaHei UI";
        using var probe = new Font(family, maxSize, style, GraphicsUnit.Pixel);
        var measured = g.MeasureString(text, probe, new SizeF(Math.Max(box.Width, 1), Math.Max(box.Height * 3, 1)));
        var widthScale = measured.Width <= 0 ? 1f : box.Width / measured.Width;
        var heightScale = measured.Height <= 0 ? 1f : box.Height / measured.Height;
        var scale = Math.Min(1f, Math.Min(widthScale, heightScale));
        var fitted = Math.Clamp((float)Math.Floor(maxSize * scale * 0.98f), minSize, maxSize);
        return new Font(family, fitted, style, GraphicsUnit.Pixel);
    }

    private static GraphicsPath RoundedRect(Rectangle rect, int radius)
    {
        var path = new GraphicsPath();
        var d = radius * 2;
        path.AddArc(rect.Left, rect.Top, d, d, 180, 90);
        path.AddArc(rect.Right - d, rect.Top, d, d, 270, 90);
        path.AddArc(rect.Right - d, rect.Bottom - d, d, d, 0, 90);
        path.AddArc(rect.Left, rect.Bottom - d, d, d, 90, 90);
        path.CloseFigure();
        return path;
    }

    private const int WM_NCLBUTTONDOWN = 0x00A1;
    private const int HTCAPTION = 2;
    private const int HTCLIENT = 1;
    private const int HTLEFT = 10;
    private const int HTRIGHT = 11;
    private const int HTTOP = 12;
    private const int HTTOPLEFT = 13;
    private const int HTTOPRIGHT = 14;
    private const int HTBOTTOM = 15;
    private const int HTBOTTOMLEFT = 16;
    private const int HTBOTTOMRIGHT = 17;
    private const int ULW_ALPHA = 0x00000002;
    private const byte AC_SRC_OVER = 0x00;
    private const byte AC_SRC_ALPHA = 0x01;

    [StructLayout(LayoutKind.Sequential)]
    private struct POINT { public int X; public int Y; public POINT(int x, int y) { X = x; Y = y; } }
    [StructLayout(LayoutKind.Sequential)]
    private struct SIZE { public int CX; public int CY; public SIZE(int cx, int cy) { CX = cx; CY = cy; } }
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    private struct BLENDFUNCTION { public byte BlendOp; public byte BlendFlags; public byte SourceConstantAlpha; public byte AlphaFormat; }

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool UpdateLayeredWindow(IntPtr hwnd, IntPtr hdcDst, ref POINT pptDst, ref SIZE psize, IntPtr hdcSrc, ref POINT pptSrc, int crKey, ref BLENDFUNCTION pblend, int dwFlags);
    [DllImport("user32.dll")]
    private static extern IntPtr GetDC(IntPtr hWnd);
    [DllImport("user32.dll")]
    private static extern int ReleaseDC(IntPtr hWnd, IntPtr hDC);
    [DllImport("gdi32.dll")]
    private static extern IntPtr CreateCompatibleDC(IntPtr hDC);
    [DllImport("gdi32.dll")]
    private static extern bool DeleteDC(IntPtr hdc);
    [DllImport("gdi32.dll")]
    private static extern IntPtr SelectObject(IntPtr hdc, IntPtr hgdiobj);
    [DllImport("gdi32.dll")]
    private static extern bool DeleteObject(IntPtr hObject);
    [DllImport("user32.dll")]
    private static extern bool ReleaseCapture();
    [DllImport("user32.dll")]
    private static extern IntPtr SendMessage(IntPtr hWnd, int Msg, IntPtr wParam, IntPtr lParam);
}

internal static class HeatmapRenderer
{
    private readonly record struct KeySpec(string Label, string Lookup, float X, float Y, float W, float H);

    private static readonly KeySpec[] KeyboardLayout = BuildKeyboardLayout();

    public static Color SurfaceColor { get; private set; } = Color.FromArgb(247, 248, 250);
    public static Color CardColor { get; private set; } = Color.White;
    public static Color BorderColor { get; private set; } = Color.FromArgb(224, 228, 233);
    public static Color TextColor { get; private set; } = Color.FromArgb(26, 32, 44);
    public static Color MutedColor { get; private set; } = Color.FromArgb(91, 103, 117);
    public static Color ColdColor { get; private set; } = Color.FromArgb(242, 245, 248);
    public static Color HotColor { get; private set; } = Color.FromArgb(230, 76, 60);
    public static Color PressColor { get; private set; } = Color.FromArgb(70, 210, 255);
    public static Func<SpeedStats>? LiveSpeedProvider { get; set; }

    public static void ApplyTheme(string theme)
    {
        switch (theme)
        {
            case "深色":
                SurfaceColor = Color.FromArgb(31, 36, 45);
                CardColor = Color.FromArgb(41, 47, 58);
                BorderColor = Color.FromArgb(82, 92, 108);
                TextColor = Color.FromArgb(235, 240, 247);
                MutedColor = Color.FromArgb(170, 181, 196);
                ColdColor = Color.FromArgb(53, 61, 74);
                HotColor = Color.FromArgb(255, 112, 88);
                PressColor = Color.FromArgb(88, 230, 255);
                break;
            case "赛博蓝":
                SurfaceColor = Color.FromArgb(8, 18, 32);
                CardColor = Color.FromArgb(12, 32, 55);
                BorderColor = Color.FromArgb(33, 117, 154);
                TextColor = Color.FromArgb(222, 247, 255);
                MutedColor = Color.FromArgb(132, 190, 212);
                ColdColor = Color.FromArgb(19, 48, 76);
                HotColor = Color.FromArgb(255, 80, 150);
                PressColor = Color.FromArgb(0, 255, 210);
                break;
            case "极简":
                SurfaceColor = Color.White;
                CardColor = Color.White;
                BorderColor = Color.FromArgb(210, 210, 210);
                TextColor = Color.FromArgb(24, 24, 24);
                MutedColor = Color.FromArgb(105, 105, 105);
                ColdColor = Color.FromArgb(246, 246, 246);
                HotColor = Color.FromArgb(38, 38, 38);
                PressColor = Color.FromArgb(80, 160, 255);
                break;
            default:
                SurfaceColor = Color.FromArgb(247, 248, 250);
                CardColor = Color.White;
                BorderColor = Color.FromArgb(224, 228, 233);
                TextColor = Color.FromArgb(26, 32, 44);
                MutedColor = Color.FromArgb(91, 103, 117);
                ColdColor = Color.FromArgb(242, 245, 248);
                HotColor = Color.FromArgb(230, 76, 60);
                PressColor = Color.FromArgb(70, 210, 255);
                break;
        }
    }



    private static KeySpec[] BuildKeyboardLayout()
    {
        var keys = new List<KeySpec>();
        void Add(string label, float x, float y, float w = 1, float h = 1, string? lookup = null) => keys.Add(new KeySpec(label, lookup ?? label, x, y, w, h));

        // 104 键全尺寸键盘，单位坐标刻意保留真实键盘分区间隔：主键区、编辑区、方向键、小数字键盘。
        Add("Esc", 0, 0);
        for (var i = 1; i <= 4; i++) Add($"F{i}", 2 + (i - 1) * 1.1f, 0);
        for (var i = 5; i <= 8; i++) Add($"F{i}", 7.1f + (i - 5) * 1.1f, 0);
        for (var i = 9; i <= 12; i++) Add($"F{i}", 12.2f + (i - 9) * 1.1f, 0);
        Add("PrtSc", 17.0f, 0, lookup: "Print Screen"); Add("ScrLk", 18.1f, 0, lookup: "Scroll Lock"); Add("Pause", 19.2f, 0);

        Add("`", 0, 1.25f); Add("1", 1.1f, 1.25f); Add("2", 2.2f, 1.25f); Add("3", 3.3f, 1.25f); Add("4", 4.4f, 1.25f); Add("5", 5.5f, 1.25f); Add("6", 6.6f, 1.25f); Add("7", 7.7f, 1.25f); Add("8", 8.8f, 1.25f); Add("9", 9.9f, 1.25f); Add("0", 11.0f, 1.25f); Add("-", 12.1f, 1.25f); Add("=", 13.2f, 1.25f); Add("Backspace", 14.3f, 1.25f, 1.7f);
        Add("Ins", 17.0f, 1.25f, lookup: "Insert"); Add("Home", 18.1f, 1.25f); Add("PgUp", 19.2f, 1.25f, lookup: "Page Up");
        Add("Num", 20.9f, 1.25f, lookup: "Num Lock"); Add("/", 22.0f, 1.25f, lookup: "Num /"); Add("*", 23.1f, 1.25f, lookup: "Num *"); Add("-", 24.2f, 1.25f, lookup: "Num -");

        Add("Tab", 0, 2.35f, 1.5f); Add("Q", 1.6f, 2.35f); Add("W", 2.7f, 2.35f); Add("E", 3.8f, 2.35f); Add("R", 4.9f, 2.35f); Add("T", 6.0f, 2.35f); Add("Y", 7.1f, 2.35f); Add("U", 8.2f, 2.35f); Add("I", 9.3f, 2.35f); Add("O", 10.4f, 2.35f); Add("P", 11.5f, 2.35f); Add("[", 12.6f, 2.35f); Add("]", 13.7f, 2.35f); Add("\\", 14.8f, 2.35f, 1.2f);
        Add("Del", 17.0f, 2.35f, lookup: "Delete"); Add("End", 18.1f, 2.35f); Add("PgDn", 19.2f, 2.35f, lookup: "Page Down");
        Add("7", 20.9f, 2.35f, lookup: "Num 7"); Add("8", 22.0f, 2.35f, lookup: "Num 8"); Add("9", 23.1f, 2.35f, lookup: "Num 9"); Add("+", 24.2f, 2.35f, 1, 2.1f, "Num +");

        Add("Caps", 0, 3.45f, 1.75f, lookup: "Caps Lock"); Add("A", 1.85f, 3.45f); Add("S", 2.95f, 3.45f); Add("D", 4.05f, 3.45f); Add("F", 5.15f, 3.45f); Add("G", 6.25f, 3.45f); Add("H", 7.35f, 3.45f); Add("J", 8.45f, 3.45f); Add("K", 9.55f, 3.45f); Add("L", 10.65f, 3.45f); Add(";", 11.75f, 3.45f); Add("'", 12.85f, 3.45f); Add("Enter", 13.95f, 3.45f, 2.05f);
        Add("4", 20.9f, 3.45f, lookup: "Num 4"); Add("5", 22.0f, 3.45f, lookup: "Num 5"); Add("6", 23.1f, 3.45f, lookup: "Num 6");

        Add("Shift", 0, 4.55f, 2.25f, lookup: "Left Shift"); Add("Z", 2.35f, 4.55f); Add("X", 3.45f, 4.55f); Add("C", 4.55f, 4.55f); Add("V", 5.65f, 4.55f); Add("B", 6.75f, 4.55f); Add("N", 7.85f, 4.55f); Add("M", 8.95f, 4.55f); Add(",", 10.05f, 4.55f); Add(".", 11.15f, 4.55f); Add("/", 12.25f, 4.55f); Add("Shift", 13.35f, 4.55f, 2.65f, lookup: "Right Shift");
        Add("↑", 18.1f, 4.55f, lookup: "Up");
        Add("1", 20.9f, 4.55f, lookup: "Num 1"); Add("2", 22.0f, 4.55f, lookup: "Num 2"); Add("3", 23.1f, 4.55f, lookup: "Num 3"); Add("Enter", 24.2f, 4.55f, 1, 2.1f, "Num Enter");

        Add("Ctrl", 0, 5.65f, 1.25f, lookup: "Left Ctrl"); Add("Win", 1.35f, 5.65f, 1.25f, lookup: "Left Win"); Add("Alt", 2.7f, 5.65f, 1.25f, lookup: "Left Alt"); Add("Space", 4.05f, 5.65f, 6.25f); Add("Alt", 10.45f, 5.65f, 1.25f, lookup: "Right Alt"); Add("Win", 11.8f, 5.65f, 1.25f, lookup: "Right Win"); Add("Menu", 13.15f, 5.65f, 1.25f, lookup: "Menu"); Add("Ctrl", 14.5f, 5.65f, 1.5f, lookup: "Right Ctrl");
        Add("←", 17.0f, 5.65f, lookup: "Left"); Add("↓", 18.1f, 5.65f, lookup: "Down"); Add("→", 19.2f, 5.65f, lookup: "Right");
        Add("0", 20.9f, 5.65f, 2.1f, lookup: "Num 0"); Add(".", 23.1f, 5.65f, lookup: "Num .");

        return keys.ToArray();
    }

    public static void DrawKeyboard(Graphics g, Rectangle bounds, UsageState state, bool title, IEnumerable<string>? pressedKeys = null)
    {
        g.SmoothingMode = SmoothingMode.AntiAlias;
        using var bg = new SolidBrush(CardColor);
        using var border = new Pen(BorderColor);
        using var text = new SolidBrush(TextColor);
        using var muted = new SolidBrush(MutedColor);
        using var titleFont = new Font("Microsoft YaHei UI", Math.Max(10F, Math.Min(20F, bounds.Height * 0.038F)), FontStyle.Bold);
        using var smallFont = new Font("Microsoft YaHei UI", Math.Max(6F, Math.Min(11F, bounds.Height * 0.024F)));

        RoundRect(g, bg, border, bounds, 8);

        // 标题、底部说明、按键尺寸都根据窗口实时缩放。
        var header = title ? Math.Clamp((int)(bounds.Height * 0.145), 58, 86) : Math.Clamp((int)(bounds.Height * 0.045), 12, 24);
        if (title)
        {
            var titleH = Math.Clamp((int)(bounds.Height * 0.075), 30, 46);
            TextRenderer.DrawText(g, "键盘热力图", titleFont, new Rectangle(bounds.X + 14, bounds.Y + 14, bounds.Width - 28, titleH), TextColor, TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis);
        }

        var max = Math.Max(1, state.Keys.Count == 0 ? 1 : state.Keys.Values.Max());
        var pressedSet = new HashSet<string>(pressedKeys ?? Array.Empty<string>(), StringComparer.OrdinalIgnoreCase);
        var marginX = Math.Clamp(bounds.Width / 90, 6, 16);
        var innerX = bounds.X + marginX;
        var innerY = bounds.Y + header;
        var innerW = bounds.Width - marginX * 2;
        var footer = Math.Clamp(bounds.Height / 9, 32, 66);
        var innerH = Math.Max(120, bounds.Height - header - footer - 6);

        // 关键改动：不再设置 9px 的最小 keySize，否则小窗口会撑爆并出现横向滚动。
        // 现在完全按窗口大小计算，窗口小则整体缩小，窗口大则整体放大。
        var keySize = Math.Min(innerW / 25.2f, innerH / 6.65f);
        keySize = Math.Max(5.2f, keySize);
        var gap = Math.Max(1.5f, keySize * 0.10f);
        var keyboardW = keySize * 25.2f;
        var keyboardH = keySize * 6.65f;
        var xOffset = innerX + Math.Max(0, (innerW - keyboardW) / 2f);
        var yOffset = innerY + Math.Max(title ? 8 : 2, (innerH - keyboardH) * 0.10f);

        foreach (var key in KeyboardLayout)
        {
            var count = state.Keys.TryGetValue(key.Lookup, out var value) ? value : 0;
            var rect = new Rectangle(
                (int)Math.Round(xOffset + key.X * keySize),
                (int)Math.Round(yOffset + key.Y * keySize),
                (int)Math.Round(key.W * keySize - gap),
                (int)Math.Round(key.H * keySize - gap));
            if (rect.Width < 7 || rect.Height < 7) continue;
            var isPressed = IsKeyPressed(key, pressedSet);
            using var fill = new SolidBrush(isPressed ? PressedFillColor(count, max) : HeatColor(count, max));
            using var keyBorder = new Pen(isPressed ? Color.FromArgb(35, 211, 255) : BorderColor, isPressed ? Math.Max(2f, rect.Height / 22f) : 1f);
            RoundRect(g, fill, keyBorder, rect, Math.Max(3, Math.Min(7, rect.Height / 8)));
            if (isPressed)
            {
                var glow = Rectangle.Inflate(rect, Math.Max(1, rect.Width / 28), Math.Max(1, rect.Height / 28));
                using var glowPen = new Pen(Color.FromArgb(0, 174, 255), Math.Max(1.5f, rect.Height / 32f));
                using var glowPath = RoundedPath(glow, Math.Max(4, Math.Min(9, glow.Height / 7)));
                g.DrawPath(glowPen, glowPath);
            }

            var pad = Math.Max(1, Math.Min(5, rect.Width / 16));
            var showCount = count > 0 && rect.Width >= 17 && rect.Height >= 17;
            var contentTop = rect.Y + Math.Max(1, rect.Height / 18);
            var contentBottom = rect.Bottom - Math.Max(1, rect.Height / 18);
            var contentHeight = Math.Max(10, contentBottom - contentTop);
            Rectangle labelRect;
            Rectangle countRect = Rectangle.Empty;
            if (showCount)
            {
                // 标签和次数分区绘制，字号按按键实际宽高二次适配，避免小窗口里挤成一团。
                var labelHeight = Math.Max(9, (int)Math.Round(contentHeight * 0.60));
                labelRect = new Rectangle(rect.X + pad, contentTop, Math.Max(1, rect.Width - pad * 2), labelHeight);
                countRect = new Rectangle(rect.X + pad, labelRect.Bottom - 1, Math.Max(1, rect.Width - pad * 2), Math.Max(8, contentBottom - labelRect.Bottom + 1));
            }
            else
            {
                labelRect = new Rectangle(rect.X + pad, contentTop, Math.Max(1, rect.Width - pad * 2), contentHeight);
            }

            var labelSize = Math.Max(8.0F, Math.Min(18.0F, Math.Min(rect.Height * 0.42F, rect.Width * 0.52F)));
            using var labelFont = new Font("Segoe UI", labelSize, FontStyle.Bold);
            DrawFittedText(g, key.Label, labelFont, TextColor, labelRect, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis | TextFormatFlags.SingleLine);
            if (showCount)
            {
                var countSize = Math.Max(7.0F, Math.Min(13.5F, Math.Min(rect.Height * 0.30F, rect.Width * 0.38F)));
                using var countFont = new Font("Segoe UI", countSize, FontStyle.Bold);
                DrawFittedText(g, count.ToString(), countFont, MutedColor, countRect, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis | TextFormatFlags.SingleLine);
            }
        }

        var unknownKeys = state.Keys.Keys.Where(k => !KeyboardLayout.Any(x => string.Equals(x.Lookup, k, StringComparison.OrdinalIgnoreCase))).OrderBy(k => k).Take(10).ToArray();
        var speed = LiveSpeedProvider?.Invoke() ?? GetSpeedSummary(state);
        var footerText = unknownKeys.Length == 0
            ? "真实 104 键布局：主键区、编辑区、方向键、小数字键盘都按分区显示。窗口较小时可横向/纵向滚动。"
            : "未放入图中的按键：" + string.Join("、", unknownKeys) + "。窗口较小时可横向/纵向滚动。";
        var footerX = bounds.X + marginX;
        var footerW = bounds.Width - marginX * 2;
        DrawFittedText(g, $"打字速度：本分钟 {speed.CurrentMinute} 次/分钟　当前小时 {speed.CurrentHour} 次/小时　平均 {speed.AverageMinute:F1} 次/分钟", smallFont, PressColor, new Rectangle(footerX, bounds.Bottom - footer, footerW, footer / 2 - 1), TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis | TextFormatFlags.SingleLine);
        DrawFittedText(g, footerText, smallFont, MutedColor, new Rectangle(footerX, bounds.Bottom - footer / 2, footerW, footer / 2 - 2), TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis | TextFormatFlags.SingleLine);
    }

    private static SpeedStats GetSpeedSummary(UsageState state)
    {
        var minutes = state.MinuteKeys is { Length: >= 1440 } ? state.MinuteKeys : new int[1440];
        var hours = state.HourlyKeys is { Length: >= 24 } ? state.HourlyKeys : new int[24];
        var now = DateTime.Now;
        var minuteIndex = now.Hour * 60 + now.Minute;
        var currentMinute = minuteIndex >= 0 && minuteIndex < minutes.Length ? minutes[minuteIndex] : 0;
        var currentHour = now.Hour >= 0 && now.Hour < hours.Length ? hours[now.Hour] : 0;
        var total = state.Keys.Values.Sum();
        var nowMinuteIndex = now.Hour * 60 + now.Minute;
        var firstMinuteIndex = Array.FindIndex(minutes, v => v > 0);
        var elapsedMinutes = firstMinuteIndex >= 0 ? Math.Max(1, nowMinuteIndex - firstMinuteIndex + 1) : 0;
        var avgMinute = elapsedMinutes > 0 ? total / (double)elapsedMinutes : 0;
        return new SpeedStats(currentMinute, currentHour, minutes.Max(), hours.Max(), minutes.Where(v => v > 0).DefaultIfEmpty(0).Average(), avgMinute, avgMinute * 60.0);
    }

    public static void DrawMouse(Graphics g, Rectangle bounds, UsageState state, IEnumerable<string>? pressedButtons = null)
    {
        g.SmoothingMode = SmoothingMode.AntiAlias;
        using var bg = new SolidBrush(CardColor);
        using var border = new Pen(BorderColor);
        using var strong = new Pen(Color.FromArgb(78, 91, 106), Math.Max(2, bounds.Width / 260));
        using var titleFont = new Font("Microsoft YaHei UI", Math.Max(14F, Math.Min(20F, bounds.Height * 0.036F)), FontStyle.Bold);
        using var labelFont = new Font("Segoe UI", Math.Max(6F, Math.Min(12F, bounds.Height * 0.022F)), FontStyle.Bold);
        using var countFont = new Font("Segoe UI", Math.Max(5.8F, Math.Min(11F, bounds.Height * 0.020F)));

        RoundRect(g, bg, border, bounds, 8);
        // 标题下移并留更大高度，避免顶部黑体字被裁掉。
        var titleRect = new Rectangle(bounds.X + 18, bounds.Y + 18, bounds.Width - 36, 50);
        TextRenderer.DrawText(g, "鼠标热力图", titleFont, titleRect, TextColor, TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis);

        var max = Math.Max(1, state.Mouse.Count == 0 ? 1 : state.Mouse.Values.Max());
        var pressedSet = new HashSet<string>(pressedButtons ?? Array.Empty<string>(), StringComparer.OrdinalIgnoreCase);
        using var liveFill = new SolidBrush(PressColor);
        using var livePen = new Pen(Color.FromArgb(0, 135, 170), Math.Max(3, bounds.Width / 180));

        var contentTop = titleRect.Bottom + 18;
        var footerH = Math.Clamp(bounds.Height / 7, 72, 96);
        var contentBottom = bounds.Bottom - footerH - 18;
        var contentW = bounds.Width - 26;
        var contentH = Math.Max(110, contentBottom - contentTop);

        // 鼠标不再特别“瘦长”，改成更自然的宽高比；仍保证在总览页中更清楚。
        var targetW = Math.Min((int)(contentW * 0.78), bounds.Width - 46);
        var targetH = (int)Math.Round(targetW * 1.42);
        if (targetH > contentH)
        {
            targetH = contentH;
            targetW = (int)Math.Round(targetH / 1.42);
        }
        if (bounds.Width < 380)
        {
            targetW = Math.Min((int)(contentW * 0.62), bounds.Width - 34);
            targetH = Math.Min(contentH, (int)Math.Round(targetW * 1.46));
        }

        var minMouseW = Math.Min(170, Math.Max(120, contentW - 84));
        var minMouseH = Math.Min(240, Math.Max(160, contentH));
        var mouseW = Math.Max(minMouseW, targetW);
        var mouseH = Math.Max(minMouseH, targetH);
        var sideWGuess = Math.Max(46, (int)(mouseW * 0.18));
        var totalVisualW = mouseW + sideWGuess + 18;
        var mouseX = bounds.X + (bounds.Width - totalVisualW) / 2 + sideWGuess + 10;
        var mouseY = contentTop + Math.Max(0, (contentH - mouseH) / 2);
        var mouseRect = new Rectangle(mouseX, mouseY, mouseW, mouseH);

        using var body = new SolidBrush(ColdColor);
        g.FillEllipse(body, mouseRect);
        g.DrawEllipse(strong, mouseRect);

        int X(float v) => mouseRect.X + (int)Math.Round(mouseRect.Width * v);
        int Y(float v) => mouseRect.Y + (int)Math.Round(mouseRect.Height * v);
        int W(float v) => (int)Math.Round(mouseRect.Width * v);
        int H(float v) => (int)Math.Round(mouseRect.Height * v);

        var leftRect = new Rectangle(X(0.08f), Y(0.06f), W(0.36f), H(0.34f));
        var rightRect = new Rectangle(X(0.56f), Y(0.06f), W(0.36f), H(0.34f));
        // 中键独立显示在滚轮下方，避免被 Wheel Up / Wheel Down 覆盖。
        var midRect = new Rectangle(X(0.42f), Y(0.38f), W(0.16f), H(0.14f));
        var wheelUpRect = new Rectangle(X(0.43f), Y(0.14f), W(0.14f), H(0.10f));
        var wheelDownRect = new Rectangle(X(0.43f), Y(0.255f), W(0.14f), H(0.10f));

        var sideW = Math.Max(46, W(0.20f));
        var sideH = Math.Max(40, H(0.16f));
        var sideX = mouseRect.X - sideW - 12;
        var fwdRect = new Rectangle(sideX, Y(0.40f), sideW, sideH);
        var backRect = new Rectangle(sideX, Y(0.60f), sideW, sideH);

        void DrawZone(string name, string label, Rectangle rect, bool small = false)
        {
            var count = state.Mouse.TryGetValue(name, out var value) ? value : 0;
            if (name == "Wheel Up" && count == 0 && state.Mouse.TryGetValue("Wheel", out var oldWheel)) count = oldWheel;
            var isPressed = pressedSet.Contains(name);
            using var fill = new SolidBrush(isPressed ? liveFill.Color : HeatColor(count, max));
            RoundRect(g, fill, isPressed ? livePen : strong, rect, Math.Max(4, rect.Height / 10));

            var labelHeight = small ? (int)(rect.Height * 0.48f) : (int)(rect.Height * 0.50f);
            var labelRect = new Rectangle(rect.X + 2, rect.Y + 2, rect.Width - 4, Math.Max(12, labelHeight - 2));
            var countRect = new Rectangle(rect.X + 2, rect.Y + labelHeight - 2, rect.Width - 4, rect.Height - labelHeight);
            DrawFittedText(g, label, labelFont, TextColor, labelRect, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis | TextFormatFlags.SingleLine);
            DrawFittedText(g, count.ToString(), countFont, MutedColor, countRect, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis | TextFormatFlags.SingleLine);
        }

        DrawZone("Left Button", "Left", leftRect);
        DrawZone("Right Button", "Right", rightRect);
        DrawZone("Wheel Up", "Up", wheelUpRect, true);
        DrawZone("Wheel Down", "Down", wheelDownRect, true);
        DrawZone("Middle Button", "Mid", midRect, true);
        DrawZone("Forward Button", "Fwd", fwdRect);
        DrawZone("Back Button", "Back", backRect);

        // 底部说明抬高并加高，避免左下角小字被裁掉。
        var footerText = $"\u9f20\u6807\u79fb\u52a8\uff1a{MainForm.FormatDistance(state.MouseDistancePixels)}\r\n\u6b21\u6570\u5df2\u663e\u793a\u5728\u6309\u952e\u4e0a\uff0c\u6eda\u8f6e\u4e0a\u4e0b\u5206\u5f00\u7edf\u8ba1\u3002";
        TextRenderer.DrawText(g, footerText, countFont, new Rectangle(bounds.X + 18, bounds.Bottom - footerH + 14, bounds.Width - 36, footerH - 24), MutedColor, TextFormatFlags.Left | TextFormatFlags.Top | TextFormatFlags.WordBreak | TextFormatFlags.WordEllipsis);
    }

    public static void DrawPeakChart(Graphics g, Rectangle bounds, UsageState state)
    {
        g.SmoothingMode = SmoothingMode.AntiAlias;
        using var bg = new SolidBrush(CardColor);
        using var border = new Pen(BorderColor);
        using var axisPen = new Pen(Color.FromArgb(150, 163, 178));
        using var gridPen = new Pen(Color.FromArgb(231, 235, 240));
        using var keyBrush = new SolidBrush(HotColor);
        using var mouseBrush = new SolidBrush(Color.FromArgb(58, 126, 194));
        using var text = new SolidBrush(TextColor);
        using var muted = new SolidBrush(MutedColor);
        using var titleFont = new Font("Microsoft YaHei UI", Math.Max(11F, Math.Min(16F, bounds.Height * 0.045F)), FontStyle.Bold);
        using var font = new Font("Microsoft YaHei UI", Math.Max(7F, Math.Min(9F, bounds.Height * 0.026F)));
        using var smallFont = new Font("Microsoft YaHei UI", Math.Max(6F, Math.Min(8F, bounds.Height * 0.022F)));

        RoundRect(g, bg, border, bounds, 8);
        TextRenderer.DrawText(g, "小时峰值图", titleFont, new Rectangle(bounds.X + 18, bounds.Y + 24, bounds.Width - 36, 42), TextColor, TextFormatFlags.Left | TextFormatFlags.VerticalCenter);
        TextRenderer.DrawText(g, "纵轴单位：次；横轴半小时为短刻度，文字只显示关键整点，避免挤在一起。旧版本数据没有小时记录，会显示为 0。", font, new Rectangle(bounds.X + 18, bounds.Y + 72, bounds.Width - 36, 26), MutedColor, TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis);

        var keys = state.HourlyKeys is { Length: >= 24 } ? state.HourlyKeys : new int[24];
        var mouse = state.HourlyMouse is { Length: >= 24 } ? state.HourlyMouse : new int[24];
        var max = Math.Max(1, Math.Max(keys.Max(), mouse.Max()));
        var niceMax = NiceAxisMax(max);

        var leftMargin = Math.Clamp(bounds.Width / 11, 54, 96);
        var rightMargin = Math.Clamp(bounds.Width / 45, 18, 32);
        var topMargin = Math.Clamp(bounds.Height / 4, 96, 150);
        var bottomMargin = Math.Clamp(bounds.Height / 5, 72, 106);
        var chart = new Rectangle(bounds.X + leftMargin, bounds.Y + topMargin, bounds.Width - leftMargin - rightMargin, bounds.Height - topMargin - bottomMargin);
        if (chart.Width < 160 || chart.Height < 100) return;

        for (var i = 0; i <= 5; i++)
        {
            var value = (int)Math.Round(niceMax * i / 5.0);
            var y = chart.Bottom - (int)Math.Round(chart.Height * i / 5.0);
            g.DrawLine(gridPen, chart.X, y, chart.Right, y);
            var label = value + " 次";
            TextRenderer.DrawText(g, label, smallFont, new Rectangle(bounds.X + 10, y - 9, leftMargin - 18, 18), Color.FromArgb(91, 103, 117), TextFormatFlags.Right | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPadding);
        }

        g.DrawLine(axisPen, chart.X, chart.Bottom, chart.Right, chart.Bottom);
        g.DrawLine(axisPen, chart.X, chart.Y, chart.X, chart.Bottom);

        var slot = chart.Width / 24f;
        var barW = Math.Max(3, (int)(slot * 0.30f));

        // 下方时间刻度：半小时只画短刻度线；文字只显示关键整点，避免糊成一排。
        // 宽窗口显示每 1 小时，普通窗口显示每 2 小时，小窗口显示每 4 小时。
        var hourLabelStep = slot >= 54 ? 1 : slot >= 28 ? 2 : 4;
        for (var half = 0; half <= 48; half++)
        {
            var xTick = chart.X + half * (slot / 2f);
            var isHour = half % 2 == 0;
            using var tickPen = new Pen(Color.FromArgb(isHour ? 150 : 80, 91, 103, 117));
            g.DrawLine(tickPen, (int)xTick, chart.Bottom, (int)xTick, chart.Bottom + (isHour ? 7 : 4));

            if (isHour)
            {
                var hour = half / 2;
                if (hour % hourLabelStep == 0)
                {
                    var label = $"{hour:00}:00";
                    TextRenderer.DrawText(
                        g,
                        label,
                        smallFont,
                        new Rectangle((int)xTick - 22, chart.Bottom + 10, 46, 18),
                        Color.FromArgb(91, 103, 117),
                        TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis | TextFormatFlags.NoPadding);
                }
            }
        }

        for (var h = 0; h < 24; h++)
        {
            var x = chart.X + (int)(h * slot + slot * 0.17f);
            var keyH = (int)Math.Round(chart.Height * (keys[h] / (double)niceMax));
            var mouseH = (int)Math.Round(chart.Height * (mouse[h] / (double)niceMax));
            if (keyH > 0) g.FillRectangle(keyBrush, x, chart.Bottom - keyH, barW, keyH);
            if (mouseH > 0) g.FillRectangle(mouseBrush, x + barW + 2, chart.Bottom - mouseH, barW, mouseH);
        }

        TextRenderer.DrawText(g, "次数", smallFont, new Rectangle(bounds.X + 16, chart.Y - 22, leftMargin - 24, 18), Color.FromArgb(91, 103, 117), TextFormatFlags.Right | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPadding);
        TextRenderer.DrawText(g, "时间", smallFont, new Rectangle(chart.Right - 44, chart.Bottom + 31, 50, 18), Color.FromArgb(91, 103, 117), TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPadding);

        g.DrawString($"峰值: {max} 次", font, text, chart.X, bounds.Bottom - 42);
        g.FillRectangle(keyBrush, chart.X + 108, bounds.Bottom - 38, 16, 10);
        g.DrawString("键盘", font, text, chart.X + 130, bounds.Bottom - 43);
        g.FillRectangle(mouseBrush, chart.X + 188, bounds.Bottom - 38, 16, 10);
        g.DrawString("鼠标", font, text, chart.X + 210, bounds.Bottom - 43);
    }

    public static void DrawSpeed(Graphics g, Rectangle bounds, UsageState state, SpeedStats liveSpeed)
    {
        g.SmoothingMode = SmoothingMode.AntiAlias;
        using var bg = new SolidBrush(CardColor);
        using var border = new Pen(BorderColor);
        using var axisPen = new Pen(MutedColor);
        using var gridPen = new Pen(Color.FromArgb(Math.Min(255, BorderColor.R + 12), Math.Min(255, BorderColor.G + 12), Math.Min(255, BorderColor.B + 12)));
        using var minuteBrush = new SolidBrush(PressColor);
        using var hourBrush = new SolidBrush(HotColor);
        using var text = new SolidBrush(TextColor);
        using var muted = new SolidBrush(MutedColor);
        using var titleFont = new Font("Microsoft YaHei UI", Math.Max(11F, Math.Min(16F, bounds.Height * 0.045F)), FontStyle.Bold);
        using var font = new Font("Microsoft YaHei UI", Math.Max(7F, Math.Min(9F, bounds.Height * 0.026F)));
        using var smallFont = new Font("Microsoft YaHei UI", Math.Max(6F, Math.Min(8F, bounds.Height * 0.022F)));

        RoundRect(g, bg, border, bounds, 8);
        TextRenderer.DrawText(g, "打字速度", titleFont, new Rectangle(bounds.X + 18, bounds.Y + 24, bounds.Width - 36, 42), TextColor, TextFormatFlags.Left | TextFormatFlags.VerticalCenter);

        var minutes = state.MinuteKeys is { Length: >= 1440 } ? state.MinuteKeys : new int[1440];
        var hours = state.HourlyKeys is { Length: >= 24 } ? state.HourlyKeys : new int[24];
        var currentMinute = liveSpeed.CurrentMinute;
        var currentHour = liveSpeed.CurrentHour;
        var maxMinute = Math.Max(liveSpeed.MaxMinute, minutes.Length == 0 ? 0 : minutes.Max());
        var maxHour = Math.Max(liveSpeed.MaxHour, hours.Length == 0 ? 0 : hours.Max());
        var avgPerMinute = minutes.Where(v => v > 0).DefaultIfEmpty(0).Average();
        var total = state.Keys.Values.Sum();

        TextRenderer.DrawText(g, $"本分钟：{currentMinute} 次/分钟　当前小时：{currentHour} 次/小时　平均：{liveSpeed.AverageMinute:F1} 次/分钟 / {liveSpeed.AverageHour:F0} 次/小时　总按键：{total} 次", font, new Rectangle(bounds.X + 18, bounds.Y + 74, bounds.Width - 36, 26), MutedColor, TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis);
        TextRenderer.DrawText(g, $"本分钟/当前小时按自然时间统计；平均速度 = 今日总按键 ÷ 从第一次按键到现在的时间；长按不重复计入。", font, new Rectangle(bounds.X + 18, bounds.Y + 102, bounds.Width - 36, 26), MutedColor, TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis);

        var leftMargin = Math.Clamp(bounds.Width / 11, 54, 94);
        var rightMargin = Math.Clamp(bounds.Width / 45, 18, 32);
        var top1 = bounds.Y + Math.Clamp(bounds.Height / 4, 130, 176);
        var chartH = Math.Max(80, (bounds.Height - Math.Clamp(bounds.Height / 3, 190, 270)) / 2);
        var minuteChart = new Rectangle(bounds.X + leftMargin, top1, bounds.Width - leftMargin - rightMargin, chartH);
        var hourChart = new Rectangle(bounds.X + leftMargin, minuteChart.Bottom + 70, bounds.Width - leftMargin - rightMargin, chartH);

        DrawBarChart(g, minuteChart, minutes, Math.Max(1, maxMinute), 60, "每分钟输入量（次/分钟）", minuteBrush, text, muted, gridPen, axisPen, smallFont);
        DrawBarChart(g, hourChart, hours, Math.Max(1, maxHour), 2, "每小时输入量（次/小时）", hourBrush, text, muted, gridPen, axisPen, smallFont);
    }

    private static void DrawBarChart(Graphics g, Rectangle chart, int[] values, int rawMax, int labelStep, string title, Brush barBrush, Brush text, Brush muted, Pen gridPen, Pen axisPen, Font font)
    {
        if (chart.Width < 120 || chart.Height < 80) return;
        var niceMax = NiceAxisMax(rawMax);
        g.DrawString(title, font, text, chart.X, chart.Y - 24);
        for (var i = 0; i <= 4; i++)
        {
            var value = (int)Math.Round(niceMax * i / 4.0);
            var y = chart.Bottom - (int)Math.Round(chart.Height * i / 4.0);
            g.DrawLine(gridPen, chart.X, y, chart.Right, y);
            TextRenderer.DrawText(g, value + " 次", font, new Rectangle(chart.X - 88, y - 9, 82, 18), MutedColor, TextFormatFlags.Right | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPadding);
        }
        g.DrawLine(axisPen, chart.X, chart.Bottom, chart.Right, chart.Bottom);
        g.DrawLine(axisPen, chart.X, chart.Y, chart.X, chart.Bottom);

        var slot = chart.Width / (float)Math.Max(1, values.Length);
        var barW = Math.Max(1, (int)Math.Floor(slot * 0.75f));
        for (var i = 0; i < values.Length; i++)
        {
            var h = (int)Math.Round(chart.Height * (values[i] / (double)niceMax));
            if (h <= 0) continue;
            var x = chart.X + (int)Math.Floor(i * slot);
            g.FillRectangle(barBrush, x, chart.Bottom - h, barW, h);
        }

        var labels = values.Length == 24 ? new[] { 0, 3, 6, 9, 12, 15, 18, 21 } : new[] { 0, 240, 480, 720, 960, 1200 };
        foreach (var idx in labels)
        {
            if (idx >= values.Length) continue;
            var x = chart.X + (int)Math.Floor(idx * slot);
            var label = values.Length == 24 ? idx.ToString("00") + "时" : (idx / 60).ToString("00") + ":00";
            TextRenderer.DrawText(g, label, font, new Rectangle(x - 10, chart.Bottom + 5, 48, 18), MutedColor, TextFormatFlags.Left | TextFormatFlags.NoPadding);
        }
        TextRenderer.DrawText(g, values.Length == 24 ? "小时" : "时间", font, new Rectangle(chart.Right - 48, chart.Bottom + 25, 54, 18), MutedColor, TextFormatFlags.Left | TextFormatFlags.NoPadding);
    }


    private static int NiceAxisMax(int value)
    {
        if (value <= 5) return 5;
        var pow = Math.Pow(10, Math.Floor(Math.Log10(value)));
        foreach (var step in new[] { 1, 2, 5, 10 })
        {
            var candidate = (int)(step * pow);
            if (candidate >= value) return candidate;
        }
        return value;
    }

    private static bool IsKeyPressed(KeySpec key, HashSet<string> pressedSet)
    {
        // 只按 Lookup 精确匹配，避免两个相同文字的键一起亮。
        // 例如主 Enter 与数字键盘 Enter、左 Shift 与右 Shift 都会分开统计和高亮。
        return pressedSet.Count > 0 && pressedSet.Contains(key.Lookup);
    }

    private static Color PressedFillColor(int count, int max)
    {
        var heat = HeatColor(count, max);
        return Lerp(heat, PressColor, 0.78);
    }

    private static Color HeatColor(int count, int max)
    {
        if (count <= 0) return ColdColor;
        var t = Math.Clamp(Math.Log(count + 1) / Math.Log(max + 1), 0, 1);
        return Lerp(Color.FromArgb(255, 236, 205), HotColor, t);
    }

    private static Color Lerp(Color a, Color b, double t) => Color.FromArgb((int)(a.R + (b.R - a.R) * t), (int)(a.G + (b.G - a.G) * t), (int)(a.B + (b.B - a.B) * t));

    private static void RoundRect(Graphics g, Brush fill, Pen pen, Rectangle rect, int radius)
    {
        using var path = RoundedPath(rect, radius);
        g.FillPath(fill, path);
        g.DrawPath(pen, path);
    }

    private static GraphicsPath RoundedPath(Rectangle rect, int radius)
    {
        var d = radius * 2;
        var path = new GraphicsPath();
        path.AddArc(rect.X, rect.Y, d, d, 180, 90);
        path.AddArc(rect.Right - d, rect.Y, d, d, 270, 90);
        path.AddArc(rect.Right - d, rect.Bottom - d, d, d, 0, 90);
        path.AddArc(rect.X, rect.Bottom - d, d, d, 90, 90);
        path.CloseFigure();
        return path;
    }

    private static void DrawFittedString(Graphics g, string value, Font font, Brush brush, Rectangle rect)
    {
        using var format = new StringFormat { Trimming = StringTrimming.EllipsisCharacter, FormatFlags = StringFormatFlags.NoWrap };
        g.DrawString(value, font, brush, rect, format);
    }

    private static void DrawFittedText(Graphics g, string value, Font font, Color color, Rectangle rect, TextFormatFlags flags)
    {
        if (rect.Width <= 0 || rect.Height <= 0 || string.IsNullOrEmpty(value)) return;

        var drawFlags = flags | TextFormatFlags.NoPadding;
        var measureFlags = drawFlags & ~TextFormatFlags.EndEllipsis & ~TextFormatFlags.WordEllipsis;
        var proposed = new Size(Math.Max(1, rect.Width), Math.Max(1, rect.Height));
        var measured = TextRenderer.MeasureText(value, font, proposed, measureFlags);
        if (measured.Width <= rect.Width && measured.Height <= rect.Height)
        {
            TextRenderer.DrawText(g, value, font, rect, color, drawFlags);
            return;
        }

        // v1.7.4 这里逐级循环 MeasureText/创建 Font，按键高亮每帧要画 100+ 个键时会明显卡顿。
        // 改成一次测量 + 一次缩放，保留自适应但避免每次 Paint 做大量循环。
        var scaleW = measured.Width <= 0 ? 1F : rect.Width / (float)measured.Width;
        var scaleH = measured.Height <= 0 ? 1F : rect.Height / (float)measured.Height;
        var scaledSize = Math.Max(5.2F, Math.Min(font.Size, font.Size * Math.Min(scaleW, scaleH) * 0.94F));
        using var fitted = new Font(font.FontFamily, scaledSize, font.Style);
        TextRenderer.DrawText(g, value, fitted, rect, color, drawFlags);
    }

    private static void DrawCentered(Graphics g, string value, Font font, Brush brush, Rectangle rect)
    {
        using var format = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center, Trimming = StringTrimming.EllipsisCharacter };
        g.DrawString(value, font, brush, rect, format);
    }
}



internal sealed class UsageState
{
    public string Date { get; set; } = "";
    public Dictionary<string, int> Keys { get; set; } = new();
    public Dictionary<string, int> Mouse { get; set; } = new();
    public int[] HourlyKeys { get; set; } = new int[24];
    public int[] HourlyMouse { get; set; } = new int[24];
    public int[] MinuteKeys { get; set; } = new int[1440];
    public Dictionary<string, int> Combos { get; set; } = new();
    public Dictionary<string, int> RapidKeys { get; set; } = new();
    public Dictionary<string, Dictionary<string, int>> AppKeys { get; set; } = new();
    public Dictionary<string, double> AppUsageSeconds { get; set; } = new();
    public Dictionary<string, string> AppPaths { get; set; } = new();
    public double MouseDistancePixels { get; set; }
}

internal sealed class AppSettings
{
    public int X { get; set; } = -32000;
    public int Y { get; set; } = -32000;
    public int Width { get; set; } = 1060;
    public int Height { get; set; } = 720;
    public bool TopMost { get; set; }
    public string TrackMode { get; set; } = "全部统计";
    public string Range { get; set; } = "今天";
    public string Theme { get; set; } = "浅色";
    public bool MinimizeToBackground { get; set; } = true;
    public bool AutoStart { get; set; } = true;
    public bool TrayIconVisible { get; set; }
    public bool OverlayVisible { get; set; }
    public int SelectedTabIndex { get; set; }
    public float ListZoom { get; set; } = 1.0f;
    public int RankSortColumn { get; set; } = 3;
    public SortOrder RankSortOrder { get; set; } = SortOrder.Descending;
    public int GroupSortColumn { get; set; } = 1;
    public SortOrder GroupSortOrder { get; set; } = SortOrder.Descending;
    public int AppSortColumn { get; set; } = 1;
    public SortOrder AppSortOrder { get; set; } = SortOrder.Descending;
    public float OverviewTopRatio { get; set; } = 0.68f;
    public float OverviewMouseRatio { get; set; } = 0.23f;
}

internal static class ForegroundApp
{
    private static ForegroundAppInfo? cachedInfo;
    private static long cachedAt;

    public static ForegroundAppInfo? GetExternalAppInfo()
    {
        var now = Environment.TickCount64;
        if (now - cachedAt < 250) return cachedInfo;
        cachedAt = now;

        if (!TryGetExternalForegroundProcess(out var pid, out var name))
            return cachedInfo = null;

        var path = TryGetProcessPath(pid);
        cachedInfo = new ForegroundAppInfo(name, path);
        return cachedInfo;
    }

    public static bool TryGetExternalForegroundProcess(out uint pid, out string processName)
    {
        pid = 0;
        processName = string.Empty;

        try
        {
            var hwnd = GetForegroundWindow();
            if (hwnd == IntPtr.Zero) return false;
            GetWindowThreadProcessId(hwnd, out pid);
            if (pid == 0 || pid == Environment.ProcessId) return false;
            using var process = Process.GetProcessById((int)pid);
            processName = process.ProcessName;
            return !string.IsNullOrWhiteSpace(processName);
        }
        catch
        {
            pid = 0;
            processName = string.Empty;
            return false;
        }
    }

    public static string? TryGetProcessPath(uint pid)
    {
        const uint processQueryLimitedInformation = 0x1000;
        IntPtr handle = IntPtr.Zero;
        try
        {
            handle = OpenProcess(processQueryLimitedInformation, false, pid);
            if (handle == IntPtr.Zero) return null;

            var capacity = 1024;
            var sb = new StringBuilder(capacity);
            var size = sb.Capacity;
            if (!QueryFullProcessImageName(handle, 0, sb, ref size)) return null;
            return size > 0 ? sb.ToString(0, size) : null;
        }
        catch
        {
            return null;
        }
        finally
        {
            if (handle != IntPtr.Zero) CloseHandle(handle);
        }
    }

    [DllImport("user32.dll")]
    private static extern IntPtr GetForegroundWindow();

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern IntPtr OpenProcess(uint desiredAccess, [MarshalAs(UnmanagedType.Bool)] bool inheritHandle, uint processId);

    [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    private static extern bool QueryFullProcessImageName(IntPtr hProcess, int flags, StringBuilder exeName, ref int size);

    [DllImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool CloseHandle(IntPtr hObject);
}

internal static class AppIconProvider
{
    public static string IconKey(string appName, string? path) =>
        !string.IsNullOrWhiteSpace(path) && File.Exists(path)
            ? "path:" + path.ToLowerInvariant()
            : "app:" + appName.ToLowerInvariant();

    public static Bitmap GetIconBitmap(string? path)
    {
        try
        {
            if (!string.IsNullOrWhiteSpace(path) && File.Exists(path))
            {
                using var icon = Icon.ExtractAssociatedIcon(path);
                if (icon != null) return icon.ToBitmap();
            }
        }
        catch { }
        return SystemIcons.Application.ToBitmap();
    }

    public static string? FindRunningProcessPath(string appName)
    {
        try
        {
            var normalized = NormalizeProcessName(appName);
            foreach (var process in Process.GetProcessesByName(normalized))
            {
                using (process)
                {
                    try
                    {
                        var path = process.MainModule?.FileName;
                        if (!string.IsNullOrWhiteSpace(path) && File.Exists(path)) return path;
                    }
                    catch { }
                }
            }
        }
        catch { }
        return null;
    }

    private static string NormalizeProcessName(string appName)
    {
        var name = appName.Trim();
        if (name.EndsWith(".exe", StringComparison.OrdinalIgnoreCase)) name = name[..^4];
        return name;
    }
}

internal static class KeyNames
{
    private static readonly Dictionary<int, string> Names = Build();

    public static string NameFromVk(int vk)
    {
        if (Names.TryGetValue(vk, out var name)) return name;
        return $"VK {vk}";
    }

    public static string NameFromKeyboardHook(int vk, int scanCode, int flags)
    {
        const int extendedFlag = 0x01;
        var extended = (flags & extendedFlag) != 0;

        return vk switch
        {
            13 => extended ? "Num Enter" : "Enter",
            16 or 160 or 161 => (vk == 161 || scanCode == 0x36 || scanCode == 54) ? "Right Shift" : "Left Shift",
            17 or 162 or 163 => (vk == 163 || extended) ? "Right Ctrl" : "Left Ctrl",
            18 or 164 or 165 => (vk == 165 || extended) ? "Right Alt" : "Left Alt",
            91 => "Left Win",
            92 => "Right Win",
            93 => "Menu",
            _ => NameFromVk(vk)
        };
    }

    private static Dictionary<int, string> Build()
    {
        var names = new Dictionary<int, string>
        {
            [8] = "Backspace", [9] = "Tab", [13] = "Enter", [16] = "Shift", [17] = "Ctrl", [18] = "Alt",
            [19] = "Pause", [20] = "Caps Lock", [27] = "Esc", [32] = "Space", [44] = "Print Screen", [145] = "Scroll Lock", [33] = "Page Up",
            [34] = "Page Down", [35] = "End", [36] = "Home", [37] = "Left", [38] = "Up", [39] = "Right",
            [40] = "Down", [45] = "Insert", [46] = "Delete", [91] = "Left Win", [92] = "Right Win", [93] = "Menu",
            [96] = "Num 0", [97] = "Num 1", [98] = "Num 2", [99] = "Num 3", [100] = "Num 4",
            [101] = "Num 5", [102] = "Num 6", [103] = "Num 7", [104] = "Num 8", [105] = "Num 9",
            [106] = "Num *", [107] = "Num +", [108] = "Num Enter", [109] = "Num -", [110] = "Num .", [111] = "Num /", [144] = "Num Lock",
            [160] = "Left Shift", [161] = "Right Shift", [162] = "Left Ctrl", [163] = "Right Ctrl", [164] = "Left Alt", [165] = "Right Alt",
            [166] = "Browser Back", [167] = "Browser Forward", [168] = "Browser Refresh", [169] = "Browser Stop",
            [170] = "Browser Search", [171] = "Browser Favorites", [172] = "Browser Home",
            [173] = "Volume Mute", [174] = "Volume Down", [175] = "Volume Up",
            [176] = "Media Next", [177] = "Media Prev", [178] = "Media Stop", [179] = "Media Play/Pause",
            [180] = "Mail", [181] = "Media Select", [182] = "App 1", [183] = "App 2",
            [186] = ";", [187] = "=", [188] = ",", [189] = "-", [190] = ".", [191] = "/", [192] = "`",
            [219] = "[", [220] = "\\", [221] = "]", [222] = "'"
        };
        for (var i = 48; i <= 57; i++) names[i] = ((char)i).ToString();
        for (var i = 65; i <= 90; i++) names[i] = ((char)i).ToString();
        for (var i = 112; i <= 123; i++) names[i] = $"F{i - 111}";
        return names;
    }
}

internal static class RunningAppUsageSnapshot
{
    private static readonly string WindowsDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Windows);
    private static readonly HashSet<string> IgnoredProcessNames = new(StringComparer.OrdinalIgnoreCase)
    {
        "Idle", "System", "Registry", "smss", "csrss", "wininit", "services", "lsass", "svchost",
        "fontdrvhost", "dwm", "conhost", "RuntimeBroker", "SearchIndexer", "SearchProtocolHost",
        "SearchFilterHost", "audiodg", "WmiPrvSE", "AggregatorHost", "SecurityHealthService"
    };

    public static IEnumerable<ForegroundAppInfo> GetRunningExternalApps()
    {
        var current = Process.GetCurrentProcess();
        var currentSessionId = current.SessionId;
        var currentProcessId = current.Id;
        var result = new Dictionary<string, ForegroundAppInfo>(StringComparer.OrdinalIgnoreCase);

        foreach (var process in Process.GetProcesses())
        {
            using (process)
            {
                try
                {
                    if (process.Id == currentProcessId) continue;
                    if (process.SessionId != currentSessionId) continue;
                    var name = process.ProcessName;
                    if (string.IsNullOrWhiteSpace(name)) continue;
                    if (IgnoredProcessNames.Contains(name)) continue;

                    string? path = null;
                    try { path = process.MainModule?.FileName; } catch { }

                    if (!ShouldCountProcess(process, path)) continue;

                    var key = name.Trim();
                    if (!result.TryGetValue(key, out var existing) || string.IsNullOrWhiteSpace(existing.ExecutablePath))
                        result[key] = new ForegroundAppInfo(name, path);
                }
                catch
                {
                    // 有些系统/权限较高进程无法读取，直接跳过，避免影响实时输入。
                }
            }
        }

        var foreground = ForegroundApp.GetExternalAppInfo();
        if (foreground is { } app && !string.IsNullOrWhiteSpace(app.Name))
            result[app.Name] = app;

        return result.Values;
    }

    private static bool ShouldCountProcess(Process process, string? path)
    {
        if (process.MainWindowHandle != IntPtr.Zero) return true;
        if (string.IsNullOrWhiteSpace(path)) return false;
        return !IsWindowsSystemPath(path);
    }

    private static bool IsWindowsSystemPath(string path)
    {
        try
        {
            var fullPath = Path.GetFullPath(path);
            if (string.IsNullOrWhiteSpace(WindowsDirectory)) return false;
            var windowsPath = Path.GetFullPath(WindowsDirectory).TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar;
            return fullPath.StartsWith(windowsPath, StringComparison.OrdinalIgnoreCase);
        }
        catch
        {
            return false;
        }
    }
}

internal static class InputStateProbe
{
    private static readonly Dictionary<string, int> KeyboardVirtualKeys = new(StringComparer.OrdinalIgnoreCase)
    {
        ["Backspace"] = 0x08, ["Tab"] = 0x09, ["Enter"] = 0x0D, ["Shift"] = 0x10, ["Ctrl"] = 0x11, ["Alt"] = 0x12, ["Esc"] = 0x1B, ["Space"] = 0x20,
        ["Page Up"] = 0x21, ["Page Down"] = 0x22, ["End"] = 0x23, ["Home"] = 0x24,
        ["Left"] = 0x25, ["Up"] = 0x26, ["Right"] = 0x27, ["Down"] = 0x28,
        ["Print Screen"] = 0x2C, ["Insert"] = 0x2D, ["Delete"] = 0x2E,
        ["Left Shift"] = 0xA0, ["Right Shift"] = 0xA1, ["Left Ctrl"] = 0xA2, ["Right Ctrl"] = 0xA3,
        ["Left Alt"] = 0xA4, ["Right Alt"] = 0xA5, ["Left Win"] = 0x5B, ["Right Win"] = 0x5C, ["Menu"] = 0x5D,
        ["Caps Lock"] = 0x14, ["Num Lock"] = 0x90, ["Scroll Lock"] = 0x91,
        ["Num 0"] = 0x60, ["Num 1"] = 0x61, ["Num 2"] = 0x62, ["Num 3"] = 0x63, ["Num 4"] = 0x64,
        ["Num 5"] = 0x65, ["Num 6"] = 0x66, ["Num 7"] = 0x67, ["Num 8"] = 0x68, ["Num 9"] = 0x69,
        ["Num *"] = 0x6A, ["Num +"] = 0x6B, ["Num Enter"] = 0x0D, ["Num -"] = 0x6D, ["Num ."] = 0x6E, ["Num /"] = 0x6F,
        [";"] = 0xBA, ["="] = 0xBB, [","] = 0xBC, ["-"] = 0xBD, ["."] = 0xBE, ["/"] = 0xBF,
        ["`"] = 0xC0, ["["] = 0xDB, ["\\"] = 0xDC, ["]"] = 0xDD, ["'"] = 0xDE,
        ["Browser Back"] = 0xA6, ["Browser Forward"] = 0xA7, ["Browser Refresh"] = 0xA8, ["Browser Stop"] = 0xA9,
        ["Browser Search"] = 0xAA, ["Browser Favorites"] = 0xAB, ["Browser Home"] = 0xAC,
        ["Volume Mute"] = 0xAD, ["Volume Down"] = 0xAE, ["Volume Up"] = 0xAF,
        ["Media Next"] = 0xB0, ["Media Prev"] = 0xB1, ["Media Stop"] = 0xB2, ["Media Play/Pause"] = 0xB3,
        ["Mail"] = 0xB4, ["Media Select"] = 0xB5, ["App 1"] = 0xB6, ["App 2"] = 0xB7
    };

    static InputStateProbe()
    {
        for (var i = 0; i <= 9; i++) KeyboardVirtualKeys[i.ToString()] = 0x30 + i;
        for (var c = 'A'; c <= 'Z'; c++) KeyboardVirtualKeys[c.ToString()] = c;
        for (var i = 1; i <= 24; i++) KeyboardVirtualKeys[$"F{i}"] = 0x6F + i;
    }

    public static bool CanProbeKeyboardKey(string name) => KeyboardVirtualKeys.ContainsKey(name);

    public static bool IsKeyboardKeyDown(string name)
    {
        if (!KeyboardVirtualKeys.TryGetValue(name, out var vk)) return false;
        return (GetAsyncKeyState(vk) & 0x8000) != 0;
    }

    public static bool IsMouseButtonDown(string name)
    {
        var vk = name switch
        {
            "Left Button" => 0x01,
            "Right Button" => 0x02,
            "Middle Button" => 0x04,
            "Back Button" => 0x05,
            "Forward Button" => 0x06,
            _ => 0
        };
        return vk != 0 && (GetAsyncKeyState(vk) & 0x8000) != 0;
    }

    [DllImport("user32.dll")]
    private static extern short GetAsyncKeyState(int vKey);
}

internal static class RawKeyboardInput
{
    public static event Action<string>? KeyPressed;
    public static event Action<string>? KeyReleased;

    private const int WmInput = 0x00FF;
    private const int RidInput = 0x10000003;
    private const int RIMTypeKeyboard = 1;
    private const int RidevInputSink = 0x00000100;
    private const int RidevRemove = 0x00000001;

    private const int WmKeydown = 0x0100;
    private const int WmKeyup = 0x0101;
    private const int WmSyskeydown = 0x0104;
    private const int WmSyskeyup = 0x0105;
    private const ushort RiKeyBreak = 0x0001;
    private const ushort RiKeyE0 = 0x0002;

    public static void Register(IntPtr hwnd)
    {
        var device = new RawInputDevice
        {
            UsagePage = 0x01,
            Usage = 0x06,
            Flags = RidevInputSink,
            Target = hwnd
        };
        RegisterRawInputDevices([device], 1, Marshal.SizeOf<RawInputDevice>());
    }

    public static void Unregister()
    {
        var device = new RawInputDevice
        {
            UsagePage = 0x01,
            Usage = 0x06,
            Flags = RidevRemove,
            Target = IntPtr.Zero
        };
        RegisterRawInputDevices([device], 1, Marshal.SizeOf<RawInputDevice>());
    }

    public static bool ProcessMessage(Message message)
    {
        if (message.Msg != WmInput) return false;

        var size = 0u;
        _ = GetRawInputData(message.LParam, RidInput, IntPtr.Zero, ref size, (uint)Marshal.SizeOf<RawInputHeader>());
        if (size == 0) return false;

        var buffer = Marshal.AllocHGlobal((int)size);
        try
        {
            if (GetRawInputData(message.LParam, RidInput, buffer, ref size, (uint)Marshal.SizeOf<RawInputHeader>()) != size)
                return false;

            var input = Marshal.PtrToStructure<RawInput>(buffer);
            if (input.Header.Type != RIMTypeKeyboard) return false;

            var keyboard = input.Keyboard;
            if (keyboard.VKey is 0 or 255) return true;

            var messageCode = (int)keyboard.Message;
            var keyFlagsForName = (keyboard.Flags & RiKeyE0) != 0 ? 0x01 : 0;
            var name = KeyNames.NameFromKeyboardHook(keyboard.VKey, keyboard.MakeCode, keyFlagsForName);

            if (messageCode == WmKeydown || messageCode == WmSyskeydown || (keyboard.Flags & RiKeyBreak) == 0)
            {
                KeyPressed?.Invoke(name);
            }
            else if (messageCode == WmKeyup || messageCode == WmSyskeyup || (keyboard.Flags & RiKeyBreak) != 0)
            {
                KeyReleased?.Invoke(name);
            }
        }
        finally
        {
            Marshal.FreeHGlobal(buffer);
        }

        return true;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct RawInputDevice
    {
        public ushort UsagePage;
        public ushort Usage;
        public int Flags;
        public IntPtr Target;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct RawInputHeader
    {
        public int Type;
        public int Size;
        public IntPtr Device;
        public IntPtr WParam;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct RawKeyboard
    {
        public ushort MakeCode;
        public ushort Flags;
        public ushort Reserved;
        public ushort VKey;
        public uint Message;
        public uint ExtraInformation;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct RawInput
    {
        public RawInputHeader Header;
        public RawKeyboard Keyboard;
    }

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool RegisterRawInputDevices(RawInputDevice[] devices, uint numDevices, int size);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern uint GetRawInputData(IntPtr rawInput, uint command, IntPtr data, ref uint size, uint sizeHeader);
}

internal static class RawMouseInput
{
    public static event Action<string>? MousePressed;
    public static event Action<string>? MouseReleased;

    private const int WmInput = 0x00FF;
    private const int RidInput = 0x10000003;
    private const int RIMTypeMouse = 0;
    private const int RidevInputSink = 0x00000100;
    private const int RidevRemove = 0x00000001;

    private const ushort LeftDown = 0x0001;
    private const ushort LeftUp = 0x0002;
    private const ushort RightDown = 0x0004;
    private const ushort RightUp = 0x0008;
    private const ushort MiddleDown = 0x0010;
    private const ushort MiddleUp = 0x0020;
    private const ushort BackDown = 0x0040;     // RI_MOUSE_BUTTON_4_DOWN / XBUTTON1
    private const ushort BackUp = 0x0080;       // RI_MOUSE_BUTTON_4_UP
    private const ushort ForwardDown = 0x0100;  // RI_MOUSE_BUTTON_5_DOWN / XBUTTON2
    private const ushort ForwardUp = 0x0200;    // RI_MOUSE_BUTTON_5_UP
    private const ushort Wheel = 0x0400;

    public static void Register(IntPtr hwnd)
    {
        var device = new RawInputDevice
        {
            UsagePage = 0x01,
            Usage = 0x02,
            Flags = RidevInputSink,
            Target = hwnd
        };
        RegisterRawInputDevices([device], 1, Marshal.SizeOf<RawInputDevice>());
    }

    public static void Unregister()
    {
        var device = new RawInputDevice
        {
            UsagePage = 0x01,
            Usage = 0x02,
            Flags = RidevRemove,
            Target = IntPtr.Zero
        };
        RegisterRawInputDevices([device], 1, Marshal.SizeOf<RawInputDevice>());
    }

    public static bool ProcessMessage(Message message)
    {
        if (message.Msg != WmInput) return false;

        var size = 0u;
        _ = GetRawInputData(message.LParam, RidInput, IntPtr.Zero, ref size, (uint)Marshal.SizeOf<RawInputHeader>());
        if (size == 0) return false;

        var buffer = Marshal.AllocHGlobal((int)size);
        try
        {
            if (GetRawInputData(message.LParam, RidInput, buffer, ref size, (uint)Marshal.SizeOf<RawInputHeader>()) != size)
                return false;

            var input = Marshal.PtrToStructure<RawInput>(buffer);
            if (input.Header.Type != RIMTypeMouse) return false;

            var flags = (ushort)(input.Mouse.Buttons & 0xffff);
            var data = unchecked((short)((input.Mouse.Buttons >> 16) & 0xffff));

            if ((flags & LeftDown) != 0) MousePressed?.Invoke("Left Button");
            if ((flags & LeftUp) != 0) MouseReleased?.Invoke("Left Button");
            if ((flags & RightDown) != 0) MousePressed?.Invoke("Right Button");
            if ((flags & RightUp) != 0) MouseReleased?.Invoke("Right Button");
            if ((flags & MiddleDown) != 0) MousePressed?.Invoke("Middle Button");
            if ((flags & MiddleUp) != 0) MouseReleased?.Invoke("Middle Button");
            // Raw Input 侧键不是一个统一的 XDown/XUp：
            // 0x0040/0x0080 是 XBUTTON1（通常为 Back），0x0100/0x0200 是 XBUTTON2（通常为 Forward）。
            // 旧版本把 0x0080 当成 Down、0x0100 当成 Up，会导致前进/后退错位或漏记。
            if ((flags & BackDown) != 0) MousePressed?.Invoke("Back Button");
            if ((flags & BackUp) != 0) MouseReleased?.Invoke("Back Button");
            if ((flags & ForwardDown) != 0) MousePressed?.Invoke("Forward Button");
            if ((flags & ForwardUp) != 0) MouseReleased?.Invoke("Forward Button");
            if ((flags & Wheel) != 0) MousePressed?.Invoke(data > 0 ? "Wheel Up" : "Wheel Down");
        }
        finally
        {
            Marshal.FreeHGlobal(buffer);
        }

        return true;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct RawInputDevice
    {
        public ushort UsagePage;
        public ushort Usage;
        public int Flags;
        public IntPtr Target;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct RawInputHeader
    {
        public int Type;
        public int Size;
        public IntPtr Device;
        public IntPtr WParam;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct RawMouse
    {
        public ushort Flags;
        public uint Buttons;
        public uint RawButtons;
        public int LastX;
        public int LastY;
        public uint ExtraInformation;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct RawInput
    {
        public RawInputHeader Header;
        public RawMouse Mouse;
    }

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool RegisterRawInputDevices(RawInputDevice[] devices, uint numDevices, int size);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern uint GetRawInputData(IntPtr rawInput, uint command, IntPtr data, ref uint size, uint sizeHeader);
}
