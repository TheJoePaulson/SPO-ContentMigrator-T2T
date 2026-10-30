using System.IO;
using System.Collections.Concurrent;
using System.ComponentModel;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;
using ContentMigrator.Config;
using ContentMigrator.Engine;
using ContentMigrator.Infrastructure;
using ContentMigrator.SharePoint;

namespace ContentMigrator.UI;

public partial class MainWindow : Window
{
    private static readonly OperationKind[] OperationByIndex =
    {
        OperationKind.Inventory, OperationKind.FullMove, OperationKind.FullCopy, OperationKind.CsvMove,
        OperationKind.CsvCopy, OperationKind.ListItemCopy, OperationKind.Retry, OperationKind.CsvTemplate,
        OperationKind.CreateFromExisting, OperationKind.ModernizeList
    };

    private static readonly HashSet<OperationKind> NeedsSource = new()
    {
        OperationKind.Inventory, OperationKind.FullMove, OperationKind.FullCopy, OperationKind.ListItemCopy,
        OperationKind.CreateFromExisting, OperationKind.ModernizeList
    };

    private static readonly HashSet<OperationKind> NeedsCsv = new()
    {
        OperationKind.CsvMove, OperationKind.CsvCopy, OperationKind.Retry
    };

    private static readonly HashSet<OperationKind> NeedsDestinationList = new()
    {
        OperationKind.FullMove, OperationKind.FullCopy, OperationKind.ListItemCopy, OperationKind.CreateFromExisting,
        OperationKind.Retry, OperationKind.ModernizeList
    };

    private static readonly HashSet<OperationKind> WriteOperations = new()
    {
        OperationKind.FullMove, OperationKind.FullCopy, OperationKind.CsvMove, OperationKind.CsvCopy,
        OperationKind.ListItemCopy, OperationKind.CreateFromExisting, OperationKind.Retry, OperationKind.ModernizeList
    };

    private static readonly HashSet<OperationKind> SameSiteGuarded = new()
    {
        OperationKind.FullMove, OperationKind.FullCopy, OperationKind.ListItemCopy, OperationKind.CreateFromExisting,
        OperationKind.ModernizeList
    };

    private readonly ConcurrentQueue<string> _logQueue = new();
    private readonly DispatcherTimer _timer;
    private AppSettings _settings;
    private SpoAuthenticator _sourceAuth;
    private SpoAuthenticator _destinationAuth;
    private CancellationTokenSource _cts;
    private Task _running;
    private ProgressInfo _latestProgress;
    private RunContext _currentRun;
    private string _lastRunFolder;
    private bool _closeWhenDone;
    private bool _startupNeedsSettings;

    public MainWindow()
    {
        InitializeComponent();
        DarkTitleBar.Attach(this);
        lblSubtitle.Text = $"SharePoint list and library migration utility  -  v{typeof(MainWindow).Assembly.GetName().Version?.ToString(3)}";

        LoadSettings(firstLoad: true);
        ApplyFieldState();

        _timer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(150) };
        _timer.Tick += (_, _) => DrainQueues();
        _timer.Start();

        Loaded += OnLoaded;
        Closing += OnClosing;
    }

    // ======================================================================= settings

    private void LoadSettings(bool firstLoad)
    {
        var result = SettingsStore.Load();
        _settings = result.Settings;
        UpdateTenantLabels();

        if (!string.IsNullOrEmpty(result.Error))
        {
            if (IsLoaded) MessageBox.Show(this, result.Error, "Settings", MessageBoxButton.OK, MessageBoxImage.Warning);
            else MessageBox.Show(result.Error, "Settings", MessageBoxButton.OK, MessageBoxImage.Warning);
        }
        if (firstLoad)
        {
            _startupNeedsSettings = result.CreatedNew
                || string.IsNullOrWhiteSpace(_settings.Connection.TenantName)
                || string.IsNullOrWhiteSpace(_settings.Connection.ClientId);
        }
        ApplyCheckboxDefaults();
    }

    private void ApplyCheckboxDefaults()
    {
        chkOverwrite.IsChecked = _settings.Copy.OverwriteExisting;
        chkPreserveMeta.IsChecked = _settings.Copy.PreserveSystemMetadata;
        chkCreateDest.IsChecked = _settings.Copy.CreateDestinationIfMissing;
        chkReviewOnly.IsChecked = _settings.Provisioning.ReviewOnly;
    }

    private void OnLoaded(object sender, RoutedEventArgs e)
    {
        if (!_startupNeedsSettings) return;
        MessageBox.Show(this,
            "Before the first run, enter your tenant name and the App (client) ID of your own Entra ID app " +
            "registration in Settings > Source tenant, and check the folders. The README explains how to register the app.\n\n" +
            $"Settings are stored in:\n{SettingsStore.SettingsPath}",
            "Welcome", MessageBoxButton.OK, MessageBoxImage.Information);
        OpenSettings();
    }

    private void BtnSettings_Click(object sender, RoutedEventArgs e) => OpenSettings();

    private void OpenSettings()
    {
        var window = new SettingsWindow(_settings) { Owner = this };
        if (window.ShowDialog() == true)
        {
            _settings = window.SavedSettings;
            UpdateTenantLabels();
            if (_running == null) ApplyCheckboxDefaults();
            AppendLog($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] [INFO] Settings saved to {SettingsStore.SettingsPath}");
        }
    }

    // ======================================================================= field state

    private OperationKind SelectedOperation =>
        OperationByIndex[Math.Clamp(cmbOperation.SelectedIndex, 0, OperationByIndex.Length - 1)];

    private void CmbOperation_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (txtDestSite == null) return; // fires during InitializeComponent
        ApplyFieldState();
    }

    private void ApplyFieldState()
    {
        var idx = cmbOperation.SelectedIndex;
        var needsDest = idx is 1 or 2 or 3 or 4 or 5 or 6 or 8 or 9;
        var needsCsv = idx is 3 or 4 or 6;
        var isTemplate = idx == 7;
        txtDestSite.IsEnabled = needsDest;
        txtDestList.IsEnabled = needsDest;
        txtDestRoot.IsEnabled = idx is 1 or 2 or 3 or 4 or 6;
        txtCsvPath.IsEnabled = needsCsv;
        btnBrowseCsv.IsEnabled = needsCsv;
        txtSourceSite.IsEnabled = !isTemplate;
        txtSourceList.IsEnabled = !isTemplate;
        chkCrossTenant.IsEnabled = needsDest;
        lblDestTenant.IsEnabled = needsDest;
    }

    private void ChkCrossTenant_Changed(object sender, RoutedEventArgs e) => UpdateTenantLabels();

    /// <summary>Shows which tenants a run will use, next to the check box and in the footer.</summary>
    private void UpdateTenantLabels()
    {
        if (lblConfig == null || _settings == null) return;
        var source = string.IsNullOrWhiteSpace(_settings.Connection.TenantName) ? "(not set)" : _settings.Connection.TenantName;
        var destTenant = _settings.Destination.TenantName;
        var cross = chkCrossTenant.IsChecked == true;

        lblDestTenant.Text = string.IsNullOrWhiteSpace(destTenant)
            ? "Set the tenant in Settings > Destination tenant"
            : cross ? $"-> {destTenant}" : $"(same tenant: {source})";
        lblDestTenant.Foreground = (System.Windows.Media.Brush)FindResource(
            cross && string.IsNullOrWhiteSpace(destTenant) ? "Brush.Danger" : cross ? "Brush.Highlight" : "Brush.TextMuted");

        var tenants = cross && !string.IsNullOrWhiteSpace(destTenant)
            ? $"Source tenant: {source}  ->  Destination tenant: {destTenant}"
            : $"Tenant: {source}";
        lblConfig.Text = $"{tenants}   |   Settings: {SettingsStore.SettingsPath}";
    }

    private static string HostOf(string url) =>
        Uri.TryCreate(url, UriKind.Absolute, out var u) ? u.Host.ToLowerInvariant() : null;

    // ======================================================================= buttons

    private void BtnBrowseCsv_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new Microsoft.Win32.OpenFileDialog
        {
            Title = "Select CSV input file",
            Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*",
            Multiselect = false
        };
        if (dialog.ShowDialog(this) == true) txtCsvPath.Text = dialog.FileName;
    }

    private void BtnOpenFolder_Click(object sender, RoutedEventArgs e)
    {
        var folder = _lastRunFolder;
        if (string.IsNullOrWhiteSpace(folder) || !Directory.Exists(folder)) folder = _settings.Paths.OutputRoot;
        if (!Directory.Exists(folder))
        {
            MessageBox.Show(this, $"The folder does not exist yet:\n{folder}", "Open Output Folder", MessageBoxButton.OK,
                MessageBoxImage.Information);
            return;
        }
        RunContext.OpenFolder(folder);
    }

    private void BtnClearLog_Click(object sender, RoutedEventArgs e) => txtLog.Clear();

    private void BtnCancel_Click(object sender, RoutedEventArgs e)
    {
        _cts?.Cancel();
        lblStatus.Text = "Cancelling...";
    }

    private void BtnGenTemplate_Click(object sender, RoutedEventArgs e) => StartRun(OperationKind.CsvTemplate);

    private void BtnRun_Click(object sender, RoutedEventArgs e) => StartRun(SelectedOperation);

    // ======================================================================= run

    private void StartRun(OperationKind op)
    {
        if (_running != null) return;

        var sourceSite = txtSourceSite.Text.Trim();
        var sourceList = txtSourceList.Text.Trim();
        var destSite = txtDestSite.Text.Trim();
        var destList = txtDestList.Text.Trim();
        var csvPath = txtCsvPath.Text.Trim();

        var errors = new List<string>();
        if (NeedsSource.Contains(op) && string.IsNullOrWhiteSpace(sourceSite)) errors.Add("Source site URL is required.");
        if ((NeedsSource.Contains(op) || op == OperationKind.Retry) && string.IsNullOrWhiteSpace(sourceList))
            errors.Add("Source list/library is required.");
        if (NeedsCsv.Contains(op) && string.IsNullOrWhiteSpace(csvPath)) errors.Add("A CSV / log file path is required for this operation.");
        if (NeedsDestinationList.Contains(op) && string.IsNullOrWhiteSpace(destList)) errors.Add("Destination list/library is required.");
        if (WriteOperations.Contains(op) && chkAdminConfirm.IsChecked != true)
            errors.Add("You must confirm Site Collection Admin rights before running a write operation.");
        var crossTenant = chkCrossTenant.IsEnabled && chkCrossTenant.IsChecked == true;
        if (op != OperationKind.CsvTemplate)
        {
            errors.AddRange(SettingsStore.Validate(_settings).Select(p => "Settings: " + p));
        }
        if (crossTenant)
        {
            if (op is OperationKind.FullMove or OperationKind.CsvMove)
            {
                errors.Add("Move is not available across tenants (nothing is ever deleted from the source tenant). " +
                           "Choose the matching COPY operation instead.");
            }
            if (op is OperationKind.FullCopy or OperationKind.ListItemCopy or OperationKind.CreateFromExisting
                    or OperationKind.ModernizeList && string.IsNullOrWhiteSpace(destSite))
            {
                errors.Add("Destination site URL is required for a cross-tenant run.");
            }
            errors.AddRange(SettingsStore.ValidateCrossTenant(_settings).Select(p => "Settings: " + p));
        }
        if (errors.Count > 0)
        {
            MessageBox.Show(this, string.Join(Environment.NewLine, errors), "Missing information", MessageBoxButton.OK,
                MessageBoxImage.Warning);
            return;
        }

        // v2.3.1: a blank destination site means "same site". Catch source == destination here
        // so the operator gets an immediate dialog instead of an error buried in the log.
        var tenant = _settings.Connection.TenantName;
        var destTenant = crossTenant ? _settings.Destination.TenantName : tenant;
        var destSiteWasBlank = string.IsNullOrWhiteSpace(destSite);
        if (destSiteWasBlank && !crossTenant) destSite = sourceSite;

        // Which SharePoint host each side is on decides whether the tenant choice makes sense.
        var sourceHost = HostOf(SpoAuthenticator.ResolveSiteUrl(sourceSite, tenant));
        var destHost = HostOf(SpoAuthenticator.ResolveSiteUrl(destSite, destTenant));
        if (NeedsSource.Contains(op) && sourceHost != null && destHost != null && !destSiteWasBlank)
        {
            if (crossTenant && sourceHost == destHost)
            {
                MessageBox.Show(this,
                    $"'Destination is in another tenant' is ticked, but both sites are on {sourceHost}.{Environment.NewLine}{Environment.NewLine}" +
                    "Untick it for a same-tenant run, or enter the destination tenant's site URL.",
                    "Same tenant", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            if (!crossTenant && sourceHost != destHost && WriteOperations.Contains(op))
            {
                var answer = MessageBox.Show(this,
                    $"The destination site is on {destHost}, but the source is on {sourceHost}.{Environment.NewLine}{Environment.NewLine}" +
                    "If the destination is in another Microsoft 365 tenant, click No, tick 'Destination is in another tenant' " +
                    "and set up Settings > Destination tenant.\n\nClick Yes only if both hosts belong to the same tenant (multi-geo).",
                    "Different SharePoint host", MessageBoxButton.YesNo, MessageBoxImage.Question);
                if (answer != MessageBoxResult.Yes) return;
            }
        }

        if (SameSiteGuarded.Contains(op) &&
            string.Equals(SpoAuthenticator.ResolveSiteUrl(destSite, destTenant), SpoAuthenticator.ResolveSiteUrl(sourceSite, tenant),
                StringComparison.OrdinalIgnoreCase) &&
            string.Equals(destList, sourceList, StringComparison.OrdinalIgnoreCase))
        {
            var msg = "The destination list/library is the same as the source on the same site.";
            if (destSiteWasBlank) msg += " The destination site URL was left blank, so this is treated as a same-site operation.";
            msg += Environment.NewLine + Environment.NewLine + "Enter a different destination list/library name, or a different destination site.";
            MessageBox.Show(this, msg, "Source and destination are the same", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        // One sign-in per tenant per app session; rebuilt only if that tenant's settings change.
        try
        {
            var sourceSignIn = _settings.Connection.ToSignIn();
            if (_sourceAuth == null || !_sourceAuth.Matches(sourceSignIn)) _sourceAuth = new SpoAuthenticator(sourceSignIn);
            if (crossTenant)
            {
                var destinationSignIn = _settings.Destination.ToSignIn(_settings.Connection);
                if (_destinationAuth == null || !_destinationAuth.Matches(destinationSignIn))
                {
                    _destinationAuth = new SpoAuthenticator(destinationSignIn);
                }
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, $"Sign-in could not be configured: {ex.Message}", "Settings", MessageBoxButton.OK,
                MessageBoxImage.Error);
            return;
        }

        var request = new OperationRequest
        {
            Operation = op,
            SourceSiteUrl = sourceSite,
            SourceList = sourceList,
            DestinationSiteUrl = destSite,
            DestinationList = destList,
            DestinationRootFolder = txtDestRoot.Text.Trim(),
            CsvPath = csvPath,
            Options = new RunOptions
            {
                OverwriteExisting = chkOverwrite.IsChecked == true,
                PreserveSystemMetadata = chkPreserveMeta.IsChecked == true,
                CreateDestinationIfMissing = chkCreateDest.IsChecked == true,
                ReviewOnly = chkReviewOnly.IsChecked == true,
                CrossTenant = crossTenant
            }
        };

        // The run works on a snapshot, so editing Settings mid-run cannot change it.
        var settings = SettingsStore.Clone(_settings);
        var run = new RunContext(settings, op.ToString(), request.Options.OverwriteExisting, request.Options.PreserveSystemMetadata);
        run.LogWritten += (_, line) => _logQueue.Enqueue(line);
        run.ProgressChanged += p => _latestProgress = p;
        _lastRunFolder = run.RunFolder;
        _currentRun = run;

        txtLog.Clear();
        pbProgress.Value = 0;
        lblStatus.Text = $"Running: {op}";
        SetRunning(true);

        _cts = new CancellationTokenSource();
        var token = _cts.Token;
        var sourceAuth = _sourceAuth;
        var destinationAuth = crossTenant ? _destinationAuth : null;
        _running = Task.Run(() => OperationRunner.Run(settings, sourceAuth, destinationAuth, request, run, token));
        _running.ContinueWith(_ => Dispatcher.BeginInvoke(new Action(OnRunFinished)), TaskScheduler.Default);
    }

    private void OnRunFinished()
    {
        DrainQueues();
        var cancelled = _cts?.IsCancellationRequested == true;
        _running = null;
        _cts?.Dispose();
        _cts = null;
        SetRunning(false);
        lblStatus.Text = cancelled ? "Cancelled" : "Done";
        if (!cancelled) pbProgress.Value = 100;
        AppendLog("Operation complete.");

        var blocking = _currentRun?.BlockingMessage;
        _currentRun = null;
        if (_closeWhenDone)
        {
            Close();
            return;
        }
        if (!string.IsNullOrEmpty(blocking))
        {
            lblStatus.Text = "Stopped: site admin required";
            pbProgress.Value = 0;
            MessageBox.Show(this, blocking, "Site Collection Admin required", MessageBoxButton.OK, MessageBoxImage.Warning);
        }
    }

    private void SetRunning(bool running)
    {
        btnRun.IsEnabled = !running;
        btnGenTemplate.IsEnabled = !running;
        btnCancel.IsEnabled = running;
        cmbOperation.IsEnabled = !running;
    }

    // ======================================================================= log + progress

    private void DrainQueues()
    {
        if (!_logQueue.IsEmpty)
        {
            var sb = new StringBuilder();
            while (_logQueue.TryDequeue(out var line)) sb.AppendLine(line);
            txtLog.AppendText(sb.ToString());
            txtLog.ScrollToEnd();
        }

        var p = _latestProgress;
        if (p != null && _running != null)
        {
            pbProgress.Value = p.Percent;
            lblStatus.Text = $"{p.Percent}%  ({p.Current}/{p.Total})";
            ToolTipService.SetToolTip(lblStatus, p.Activity);
        }
    }

    private void AppendLog(string line)
    {
        txtLog.AppendText(line + Environment.NewLine);
        txtLog.ScrollToEnd();
    }

    // ======================================================================= closing

    private void OnClosing(object sender, CancelEventArgs e)
    {
        if (_running == null)
        {
            _timer.Stop();
            return;
        }
        if (_closeWhenDone)
        {
            e.Cancel = true; // already cancelling; close happens when the run stops
            return;
        }

        var answer = MessageBox.Show(this,
            "An operation is still running.\n\nCancel it and close when it stops? Logs for the work already done will be saved.",
            "Content Migrator T2T - Community Edition", MessageBoxButton.YesNo, MessageBoxImage.Question);
        e.Cancel = true;
        if (answer != MessageBoxResult.Yes) return;

        _closeWhenDone = true;
        _cts?.Cancel();
        lblStatus.Text = "Cancelling...";
    }
}
