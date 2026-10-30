namespace ContentMigrator.Infrastructure;

// Column order (= property declaration order) matches the v2.3.1 CSV logs, so existing
// spreadsheets and the Retry operation keep working. Two columns were renamed because
// their source no longer exists: PowerShellVersion -> RuntimeVersion and
// PnPModuleVersion -> CsomVersion.

public enum MigratorLogLevel { Info, Warn, Error, Success, Debug }

public sealed class OperationLogRecord
{
    public string Timestamp { get; set; }
    public string ScriptVersion { get; set; }
    public string RunId { get; set; }
    public string OperatorUpn { get; set; }
    public string MachineName { get; set; }
    public string RuntimeVersion { get; set; }
    public string CsomVersion { get; set; }
    public string Mode { get; set; }
    public string Operation { get; set; }
    public string SourceSiteUrl { get; set; }
    public string SourceServerRelativeUrl { get; set; }
    public string DestinationSiteUrl { get; set; }
    public string DestinationServerRelativeUrl { get; set; }
    public string ItemType { get; set; }
    public string Status { get; set; }
    public string ErrorMessage { get; set; }
    public string MetadataStatus { get; set; }
    public double ElapsedSeconds { get; set; }
}

public sealed class ErrorLogRecord
{
    public string Timestamp { get; set; }
    public string ScriptVersion { get; set; }
    public string RunId { get; set; }
    public string OperatorUpn { get; set; }
    public string MachineName { get; set; }
    public string Mode { get; set; }
    public string Operation { get; set; }
    public string Source { get; set; }
    public string Destination { get; set; }
    public string ErrorMessage { get; set; }
    public string ExceptionType { get; set; }
}

public sealed class MetadataLogRecord
{
    public string Timestamp { get; set; }
    public string ScriptVersion { get; set; }
    public string RunId { get; set; }
    public string OperatorUpn { get; set; }
    public string SourceServerRelativeUrl { get; set; }
    public string DestinationServerRelativeUrl { get; set; }
    public string FieldInternalName { get; set; }
    public string FieldDisplayName { get; set; }
    public string FieldType { get; set; }
    public string SourceValue { get; set; }
    public string TargetValue { get; set; }
    public string Action { get; set; }
    public string Status { get; set; }
    public string Notes { get; set; }
}

public sealed class FolderLogRecord
{
    public string Timestamp { get; set; }
    public string ScriptVersion { get; set; }
    public string RunId { get; set; }
    public string OperatorUpn { get; set; }
    public string SiteUrl { get; set; }
    public string FolderServerRelativeUrl { get; set; }
    public string Status { get; set; }
    public string Notes { get; set; }
}

public sealed class UnresolvedUserLogRecord
{
    public string Timestamp { get; set; }
    public string ScriptVersion { get; set; }
    public string RunId { get; set; }
    public string OperatorUpn { get; set; }
    public string SourceServerRelativeUrl { get; set; }
    public string FieldInternalName { get; set; }
    public string UserValue { get; set; }
    public string Notes { get; set; }
}

public sealed class ProvisionLogRecord
{
    public string Timestamp { get; set; }
    public string ScriptVersion { get; set; }
    public string RunId { get; set; }
    public string OperatorUpn { get; set; }
    public string SourceSiteUrl { get; set; }
    public string SourceList { get; set; }
    public string DestinationSiteUrl { get; set; }
    public string DestinationList { get; set; }
    public string Method { get; set; }
    public string Handlers { get; set; }
    public string Action { get; set; }
    public string Status { get; set; }
    public string TemplatePath { get; set; }
    public string Notes { get; set; }
}

public sealed class AttachmentLogRecord
{
    public string Timestamp { get; set; }
    public string ScriptVersion { get; set; }
    public string RunId { get; set; }
    public string OperatorUpn { get; set; }
    public string SourceSiteUrl { get; set; }
    public string SourceList { get; set; }
    public int SourceItemId { get; set; }
    public string DestinationSiteUrl { get; set; }
    public string DestinationList { get; set; }
    public int DestinationItemId { get; set; }
    public string FileName { get; set; }
    public double FileSizeBytes { get; set; }
    public string Status { get; set; }
    public string Notes { get; set; }
}

/// <summary>One row per distinct source person in a cross-tenant run.</summary>
public sealed class UserMappingLogRecord
{
    public string Timestamp { get; set; }
    public string ScriptVersion { get; set; }
    public string RunId { get; set; }
    public string SourceUser { get; set; }
    public string DestinationUser { get; set; }
    public string Method { get; set; }
    public string Status { get; set; }
    public string Notes { get; set; }
}

public sealed class FieldCompatibilityRecord
{
    public string ContextLabel { get; set; }
    public string FieldInternalName { get; set; }
    public string FieldDisplayName { get; set; }
    public string SourceTypeAsString { get; set; }
    public string DestinationTypeAsString { get; set; }
    public string Status { get; set; }
    public string Notes { get; set; }
}

public sealed class RunMetadataRecord
{
    public string ScriptName { get; set; }
    public string ScriptVersion { get; set; }
    public string RunId { get; set; }
    public DateTime RunStart { get; set; }
    public string RunStamp { get; set; }
    public string RunFolder { get; set; }
    public string LogFolder { get; set; }
    public string ReportFolder { get; set; }
    public string RunMode { get; set; }
    public string OperatorWindowsIdentity { get; set; }
    public string OperatorUpn { get; set; }
    public bool CrossTenant { get; set; }
    public string SourceTenant { get; set; }
    public string DestinationTenant { get; set; }
    public string DestinationOperatorUpn { get; set; }
    public string MachineName { get; set; }
    public string RuntimeVersion { get; set; }
    public string CsomVersion { get; set; }
    public string OutputRoot { get; set; }
    public bool OverwriteExisting { get; set; }
    public bool PreserveSystemMetadata { get; set; }
}

/// <summary>
/// One row of an inventory, of a Mode 4/5 CSV, or of a retry input file.
/// Inventory exports use <see cref="InventoryRecord"/>; this is the subset the copy engine needs.
/// </summary>
public sealed class ContentRow
{
    public string SourceSiteUrl { get; set; }
    public string SourceLibrary { get; set; }
    public string SourceServerRelativeUrl { get; set; }
    public string SourceRelativePathUnderRoot { get; set; }
    public string ItemType { get; set; }
    public string FileName { get; set; }
    public string DestinationSiteUrl { get; set; }
    public string DestinationLibrary { get; set; }
    public string DestinationRootFolder { get; set; }
}

public sealed class InventoryRecord
{
    public string SourceSiteUrl { get; set; }
    public string SourceLibrary { get; set; }
    public string SourceLibraryRoot { get; set; }
    public int ItemId { get; set; }
    public string UniqueId { get; set; }
    public string ItemType { get; set; }
    public string FileName { get; set; }
    public string FileExtension { get; set; }
    public string SourceServerRelativeUrl { get; set; }
    public string SourceRelativePathUnderRoot { get; set; }
    public string SourceFolderRelativePath { get; set; }
    public string SourceFolderServerRelativeUrl { get; set; }
    public string FileSizeBytes { get; set; }
    public DateTime? Created { get; set; }
    public DateTime? Modified { get; set; }
    public string AuthorEmail { get; set; }
    public string EditorEmail { get; set; }
    public string ContentType { get; set; }
    public string CheckoutUser { get; set; }

    public ContentRow ToContentRow() => new()
    {
        SourceSiteUrl = SourceSiteUrl,
        SourceLibrary = SourceLibrary,
        SourceServerRelativeUrl = SourceServerRelativeUrl,
        SourceRelativePathUnderRoot = SourceRelativePathUnderRoot,
        ItemType = ItemType,
        FileName = FileName
    };
}

public sealed class ProgressInfo
{
    public string Activity { get; init; }
    public int Current { get; init; }
    public int Total { get; init; }
    public int Percent { get; init; }
    public string Status { get; init; }
}
