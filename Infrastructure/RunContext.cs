using System.IO;
using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Text;
using System.Text.Json;
using ContentMigrator.Config;

namespace ContentMigrator.Infrastructure;

/// <summary>
/// State for one run (the former $Script: globals): identity, folders, the seven in-memory
/// logs, and their export. Log lines and progress are raised as events for the UI.
/// </summary>
public sealed class RunContext
{
    public const string AppName = "Content-Migrator-T2T-CE";

    private readonly object _logFileLock = new();
    private readonly bool _writeRunLogFile;
    private bool _runLogFileFailed;

    public RunContext(AppSettings settings, string operationName, bool overwriteExisting, bool preserveSystemMetadata)
    {
        AppVersion = Assembly.GetExecutingAssembly().GetName().Version?.ToString(3) ?? "3.0.0";
        RunStart = DateTime.Now;
        RunStamp = RunStart.ToString("yyyyMMdd_HHmmss");
        RunId = Guid.NewGuid().ToString();
        var runName = $"{RunStamp}_{SafeFileName(operationName)}";

        OutputRoot = settings.Paths.OutputRoot;
        RunFolder = Path.Combine(settings.Paths.OutputRoot, runName);
        LogFolder = Path.Combine(Blank(settings.Paths.LogRoot) ? settings.Paths.OutputRoot : settings.Paths.LogRoot, runName);
        ReportFolder = Path.Combine(Blank(settings.Paths.ReportRoot) ? settings.Paths.OutputRoot : settings.Paths.ReportRoot, runName);
        _writeRunLogFile = settings.Paths.WriteRunLogFile;

        OverwriteExisting = overwriteExisting;
        PreserveSystemMetadata = preserveSystemMetadata;

        try { OperatorWindowsIdentity = WindowsIdentity.GetCurrent().Name; }
        catch { OperatorWindowsIdentity = Environment.UserDomainName + "\\" + Environment.UserName; }
        MachineName = Environment.MachineName;
        RuntimeVersion = RuntimeInformation.FrameworkDescription;
        CsomVersion = typeof(Microsoft.SharePoint.Client.ClientContext).Assembly
            .GetCustomAttribute<AssemblyFileVersionAttribute>()?.Version
            ?? typeof(Microsoft.SharePoint.Client.ClientContext).Assembly.GetName().Version?.ToString()
            ?? "Unknown";
    }

    public string AppVersion { get; }
    public DateTime RunStart { get; }
    public string RunStamp { get; }
    public string RunId { get; }
    public string RunMode { get; } = "GUI";
    public string OutputRoot { get; }
    public string RunFolder { get; }
    public string LogFolder { get; }
    public string ReportFolder { get; }
    public string OperatorWindowsIdentity { get; }
    public string OperatorUpn { get; set; } = "Unknown";

    /// <summary>The account used in the destination tenant (cross-tenant runs only).</summary>
    public string DestinationOperatorUpn { get; set; } = "";
    public bool CrossTenant { get; set; }
    public string SourceTenant { get; set; } = "";
    public string DestinationTenant { get; set; } = "";
    public string MachineName { get; }
    public string RuntimeVersion { get; }
    public string CsomVersion { get; }
    public bool OverwriteExisting { get; }
    public bool PreserveSystemMetadata { get; }

    /// <summary>Set when the run stopped for a reason the operator must act on; the window shows it in a dialog.</summary>
    public string BlockingMessage { get; set; }

    public List<OperationLogRecord> OperationLog { get; } = new();
    public List<ErrorLogRecord> ErrorLog { get; } = new();
    public List<MetadataLogRecord> MetadataLog { get; } = new();
    public List<FolderLogRecord> FolderLog { get; } = new();
    public List<UnresolvedUserLogRecord> UnresolvedUsersLog { get; } = new();
    public List<ProvisionLogRecord> ProvisionLog { get; } = new();
    public List<AttachmentLogRecord> AttachmentLog { get; } = new();
    public List<UserMappingLogRecord> UserMappingLog { get; } = new();

    /// <summary>Raised for every log line (already formatted with timestamp and level).</summary>
    public event Action<MigratorLogLevel, string> LogWritten;

    public event Action<ProgressInfo> ProgressChanged;

    private static bool Blank(string s) => string.IsNullOrWhiteSpace(s);
    private static string Now() => DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

    public static string SafeFileName(string value) =>
        System.Text.RegularExpressions.Regex.Replace(value ?? "", @"[\\/:\*\?""<>\|]", "_");

    public void InitializeFolders()
    {
        Directory.CreateDirectory(RunFolder);
        Directory.CreateDirectory(LogFolder);
        Directory.CreateDirectory(ReportFolder);
        Log($"Output folder: {RunFolder}");
        if (!string.Equals(LogFolder, RunFolder, StringComparison.OrdinalIgnoreCase)) Log($"Log folder: {LogFolder}");
        if (!string.Equals(ReportFolder, RunFolder, StringComparison.OrdinalIgnoreCase)) Log($"Report folder: {ReportFolder}");
    }

    // ---------------------------------------------------------------- console / progress

    public void Log(string message, MigratorLogLevel level = MigratorLogLevel.Info)
    {
        var line = $"[{Now()}] [{level.ToString().ToUpperInvariant()}] {message}";
        LogWritten?.Invoke(level, line);

        if (!_writeRunLogFile || _runLogFileFailed) return;
        lock (_logFileLock)
        {
            try
            {
                Directory.CreateDirectory(LogFolder);
                File.AppendAllText(Path.Combine(LogFolder, "RunLog.txt"), line + Environment.NewLine, Encoding.UTF8);
            }
            catch
            {
                _runLogFileFailed = true; // never let the text log stop a migration
            }
        }
    }

    public void Progress(string activity, int current, int total, string currentItemName = "")
    {
        if (total <= 0) return;
        var percent = (int)Math.Round(current * 100.0 / total, 0);
        var status = string.IsNullOrWhiteSpace(currentItemName)
            ? $"Completed operation on item {current}/{total}"
            : $"Completed operation on item {current}/{total} - {currentItemName}";
        ProgressChanged?.Invoke(new ProgressInfo
        {
            Activity = activity, Current = current, Total = total, Percent = percent, Status = status
        });
    }

    // ---------------------------------------------------------------- log rows

    public void AddOperationLog(string mode, string operation, string sourceSiteUrl, string sourceServerRelativeUrl,
        string destinationSiteUrl, string destinationServerRelativeUrl, string itemType, string status,
        string errorMessage, string metadataStatus, double elapsedSeconds)
    {
        OperationLog.Add(new OperationLogRecord
        {
            Timestamp = Now(), ScriptVersion = AppVersion, RunId = RunId, OperatorUpn = OperatorUpn,
            MachineName = MachineName, RuntimeVersion = RuntimeVersion, CsomVersion = CsomVersion,
            Mode = mode, Operation = operation, SourceSiteUrl = sourceSiteUrl,
            SourceServerRelativeUrl = sourceServerRelativeUrl, DestinationSiteUrl = destinationSiteUrl,
            DestinationServerRelativeUrl = destinationServerRelativeUrl, ItemType = itemType, Status = status,
            ErrorMessage = errorMessage, MetadataStatus = metadataStatus, ElapsedSeconds = Math.Round(elapsedSeconds, 3)
        });
    }

    public void AddErrorLog(string mode, string operation, string source, string destination, string errorMessage,
        string exceptionType)
    {
        ErrorLog.Add(new ErrorLogRecord
        {
            Timestamp = Now(), ScriptVersion = AppVersion, RunId = RunId, OperatorUpn = OperatorUpn,
            MachineName = MachineName, Mode = mode, Operation = operation, Source = source,
            Destination = destination, ErrorMessage = errorMessage, ExceptionType = exceptionType
        });
    }

    public void AddErrorLog(string mode, string operation, string source, string destination, Exception ex) =>
        AddErrorLog(mode, operation, source, destination, ex.Message, ex.GetType().FullName);

    public void AddMetadataLog(string sourceServerRelativeUrl, string destinationServerRelativeUrl,
        string fieldInternalName, string fieldDisplayName, string fieldType, string sourceValue, string targetValue,
        string action, string status, string notes)
    {
        MetadataLog.Add(new MetadataLogRecord
        {
            Timestamp = Now(), ScriptVersion = AppVersion, RunId = RunId, OperatorUpn = OperatorUpn,
            SourceServerRelativeUrl = sourceServerRelativeUrl, DestinationServerRelativeUrl = destinationServerRelativeUrl,
            FieldInternalName = fieldInternalName, FieldDisplayName = fieldDisplayName, FieldType = fieldType,
            SourceValue = sourceValue, TargetValue = targetValue, Action = action, Status = status, Notes = notes
        });
    }

    public void AddFolderLog(string siteUrl, string folderServerRelativeUrl, string status, string notes)
    {
        FolderLog.Add(new FolderLogRecord
        {
            Timestamp = Now(), ScriptVersion = AppVersion, RunId = RunId, OperatorUpn = OperatorUpn,
            SiteUrl = siteUrl, FolderServerRelativeUrl = folderServerRelativeUrl, Status = status, Notes = notes
        });
    }

    public void AddUnresolvedUserLog(string sourceServerRelativeUrl, string fieldInternalName, string userValue, string notes)
    {
        UnresolvedUsersLog.Add(new UnresolvedUserLogRecord
        {
            Timestamp = Now(), ScriptVersion = AppVersion, RunId = RunId, OperatorUpn = OperatorUpn,
            SourceServerRelativeUrl = sourceServerRelativeUrl, FieldInternalName = fieldInternalName,
            UserValue = userValue, Notes = notes
        });
    }

    public void AddProvisionLog(string sourceSiteUrl, string sourceList, string destinationSiteUrl,
        string destinationList, string method, string handlers, string action, string status, string templatePath,
        string notes)
    {
        ProvisionLog.Add(new ProvisionLogRecord
        {
            Timestamp = Now(), ScriptVersion = AppVersion, RunId = RunId, OperatorUpn = OperatorUpn,
            SourceSiteUrl = sourceSiteUrl, SourceList = sourceList, DestinationSiteUrl = destinationSiteUrl,
            DestinationList = destinationList, Method = method, Handlers = handlers, Action = action,
            Status = status, TemplatePath = templatePath, Notes = notes
        });
    }

    public void AddAttachmentLog(string sourceSiteUrl, string sourceList, int sourceItemId, string destinationSiteUrl,
        string destinationList, int destinationItemId, string fileName, double fileSizeBytes, string status, string notes)
    {
        AttachmentLog.Add(new AttachmentLogRecord
        {
            Timestamp = Now(), ScriptVersion = AppVersion, RunId = RunId, OperatorUpn = OperatorUpn,
            SourceSiteUrl = sourceSiteUrl, SourceList = sourceList, SourceItemId = sourceItemId,
            DestinationSiteUrl = destinationSiteUrl, DestinationList = destinationList,
            DestinationItemId = destinationItemId, FileName = fileName, FileSizeBytes = fileSizeBytes,
            Status = status, Notes = notes
        });
    }

    public void AddUserMappingLog(string sourceUser, string destinationUser, string method, string status, string notes)
    {
        UserMappingLog.Add(new UserMappingLogRecord
        {
            Timestamp = Now(), ScriptVersion = AppVersion, RunId = RunId, SourceUser = sourceUser,
            DestinationUser = destinationUser, Method = method, Status = status, Notes = notes
        });
    }

    // ---------------------------------------------------------------- export

    public string ReportPath(string fileName) => Path.Combine(ReportFolder, fileName);
    public string RunPath(string fileName) => Path.Combine(RunFolder, fileName);

    public void ExportLogs()
    {
        Directory.CreateDirectory(LogFolder);
        Csv.Write(Path.Combine(LogFolder, "OperationLog.csv"), OperationLog);
        Csv.Write(Path.Combine(LogFolder, "ErrorLog.csv"), ErrorLog);
        Csv.Write(Path.Combine(LogFolder, "MetadataLog.csv"), MetadataLog);
        Csv.Write(Path.Combine(LogFolder, "FolderCreationLog.csv"), FolderLog);
        Csv.Write(Path.Combine(LogFolder, "UnresolvedUsers.csv"), UnresolvedUsersLog);
        Csv.Write(Path.Combine(LogFolder, "ProvisioningLog.csv"), ProvisionLog);
        Csv.Write(Path.Combine(LogFolder, "AttachmentLog.csv"), AttachmentLog);
        Csv.Write(Path.Combine(LogFolder, "UserMappingLog.csv"), UserMappingLog);
        ExportRunMetadata();
    }

    private RunMetadataRecord GetRunMetadata() => new()
    {
        ScriptName = AppName, ScriptVersion = AppVersion, RunId = RunId, RunStart = RunStart, RunStamp = RunStamp,
        RunFolder = RunFolder, LogFolder = LogFolder, ReportFolder = ReportFolder, RunMode = RunMode,
        OperatorWindowsIdentity = OperatorWindowsIdentity, OperatorUpn = OperatorUpn, CrossTenant = CrossTenant,
        SourceTenant = SourceTenant, DestinationTenant = DestinationTenant, DestinationOperatorUpn = DestinationOperatorUpn,
        MachineName = MachineName,
        RuntimeVersion = RuntimeVersion, CsomVersion = CsomVersion, OutputRoot = OutputRoot,
        OverwriteExisting = OverwriteExisting, PreserveSystemMetadata = PreserveSystemMetadata
    };

    private void ExportRunMetadata()
    {
        var metadata = GetRunMetadata();
        Csv.Write(Path.Combine(LogFolder, "RunMetadata.csv"), new[] { metadata });
        File.WriteAllText(Path.Combine(LogFolder, "RunMetadata.json"),
            JsonSerializer.Serialize(metadata, new JsonSerializerOptions { WriteIndented = true }), Encoding.UTF8);
    }

    public void ExportRunSummary()
    {
        var end = DateTime.Now;
        var elapsed = end - RunStart;
        int Ops(string status) => OperationLog.Count(o => o.Status == status);
        int Prov(Func<ProvisionLogRecord, bool> p) => ProvisionLog.Count(p);
        int Att(string status) => AttachmentLog.Count(a => a.Status == status);

        var sb = new StringBuilder();
        sb.AppendLine("Content Migrator T2T - Community Edition Run Summary");
        sb.AppendLine("================================");
        sb.AppendLine("Script metadata:");
        sb.AppendLine($"Script name: {AppName}");
        sb.AppendLine($"Version: {AppVersion}");
        sb.AppendLine($"Run ID: {RunId}");
        sb.AppendLine($"Run mode: {RunMode}");
        sb.AppendLine($"Operator: {OperatorUpn}");
        sb.AppendLine($"Source tenant: {SourceTenant}");
        if (CrossTenant)
        {
            sb.AppendLine($"Destination tenant: {DestinationTenant} (CROSS-TENANT, latest version only)");
            sb.AppendLine($"Destination operator: {DestinationOperatorUpn}");
        }
        sb.AppendLine($"Windows user: {OperatorWindowsIdentity}");
        sb.AppendLine($"Machine: {MachineName}");
        sb.AppendLine($"Runtime: {RuntimeVersion}");
        sb.AppendLine($"SharePoint CSOM: {CsomVersion}");
        sb.AppendLine($"Run started: {RunStart:yyyy-MM-dd HH:mm:ss}");
        sb.AppendLine($"Run ended: {end:yyyy-MM-dd HH:mm:ss}");
        sb.AppendLine($"Elapsed: {elapsed}");
        sb.AppendLine();
        sb.AppendLine("Output folder:");
        sb.AppendLine(RunFolder);
        sb.AppendLine("Log folder:");
        sb.AppendLine(LogFolder);
        sb.AppendLine("Report folder:");
        sb.AppendLine(ReportFolder);
        sb.AppendLine();
        sb.AppendLine("Operation totals:");
        sb.AppendLine($"Success: {Ops("Success")}");
        sb.AppendLine($"Warning: {Ops("Warning")}");
        sb.AppendLine($"Skipped: {Ops("Skipped")}");
        sb.AppendLine($"Error: {Ops("Error")}");
        sb.AppendLine();
        sb.AppendLine("Provisioning totals (create-from-existing):");
        sb.AppendLine($"Created: {Prov(p => p.Status is "Created" or "CreatedWithWarnings")}");
        sb.AppendLine($"Skipped: {Prov(p => p.Status == "Skipped")}");
        sb.AppendLine($"Error: {Prov(p => p.Status == "Error")}");
        sb.AppendLine($"Columns skipped (unsupported for provisioning): {Prov(p => p.Action == "SkipColumn")}");
        sb.AppendLine();
        if (CrossTenant)
        {
            int Map(string status) => UserMappingLog.Count(u => u.Status == status);
            sb.AppendLine("User mapping (distinct people):");
            sb.AppendLine($"Resolved: {Map("Resolved")}");
            sb.AppendLine($"Fallback account used: {Map("Fallback")}");
            sb.AppendLine($"Unresolved: {Map("Unresolved")}");
            sb.AppendLine();
        }
        sb.AppendLine("Attachments:");
        sb.AppendLine($"Copied: {Att("Copied")}");
        sb.AppendLine($"Skipped: {Att("Skipped")}");
        sb.AppendLine($"Failed: {Att("Failed")}");
        sb.AppendLine();
        sb.AppendLine("View replication:");
        sb.AppendLine($"Views created: {Prov(p => p.Action == "CreateView" && p.Status is "Created" or "CreatedWithOmittedFields")}");
        sb.AppendLine($"  ...of which had columns omitted: {Prov(p => p.Action == "CreateView" && p.Status == "CreatedWithOmittedFields")}");
        sb.AppendLine($"Views failed: {Prov(p => p.Action == "CreateView" && p.Status == "Failed")}");
        sb.AppendLine();
        sb.AppendLine("Detailed logs:");
        foreach (var f in new[] { "OperationLog.csv", "ErrorLog.csv", "MetadataLog.csv", "FolderCreationLog.csv",
                     "UnresolvedUsers.csv", "ProvisioningLog.csv", "AttachmentLog.csv", "UserMappingLog.csv", "RunLog.txt" })
        {
            sb.AppendLine(f);
        }

        File.WriteAllText(Path.Combine(LogFolder, "RunSummary.txt"), sb.ToString(), Encoding.UTF8);
    }

    public void Complete()
    {
        ExportLogs();
        ExportRunSummary();
        Log($"Logs exported to: {LogFolder}", MigratorLogLevel.Success);
    }

    public static void OpenFolder(string path)
    {
        if (string.IsNullOrWhiteSpace(path) || !Directory.Exists(path)) return;
        Process.Start(new ProcessStartInfo("explorer.exe", $"\"{path}\"") { UseShellExecute = true });
    }
}
