using System.Text.Json.Serialization;

namespace ContentMigrator.Config;

// ============================================================================
// Every value that lived in the CONFIGURATION block of the PowerShell script is
// here, grouped into sections. Each section becomes one tab in the Settings
// window (built by reflection from the attributes below), and one object in
// appsettings.json next to the .exe.
// ============================================================================

/// <summary>Marks a section class: becomes a Settings tab.</summary>
[AttributeUsage(AttributeTargets.Property)]
public sealed class SettingsSectionAttribute : Attribute
{
    public SettingsSectionAttribute(string title) { Title = title; }
    public string Title { get; }
}

/// <summary>Label + help text shown in the Settings window.</summary>
[AttributeUsage(AttributeTargets.Property)]
public sealed class SettingAttribute : Attribute
{
    public SettingAttribute(string label, string description = "")
    {
        Label = label;
        Description = description;
    }
    public string Label { get; }
    public string Description { get; }
    public int Min { get; set; } = int.MinValue;
    public int Max { get; set; } = int.MaxValue;
    public bool Required { get; set; }
    /// <summary>Show a folder Browse button next to the text box.</summary>
    public bool IsFolder { get; set; }
    /// <summary>Show a CSV file Browse button next to the text box.</summary>
    public bool IsFile { get; set; }
}

public sealed class AppSettings
{
    // JSON key stays "Connection" so settings files from 1.0 keep working.
    [SettingsSection("Source tenant")]
    public ConnectionSettings Connection { get; set; } = new();

    [SettingsSection("Destination tenant")]
    public DestinationTenantSettings Destination { get; set; } = new();

    [SettingsSection("Cross-tenant")]
    public CrossTenantSettings CrossTenant { get; set; } = new();

    [SettingsSection("Folders")]
    public PathSettings Paths { get; set; } = new();

    [SettingsSection("Copy / Move")]
    public CopySettings Copy { get; set; } = new();

    [SettingsSection("Reports")]
    public ReportSettings Reports { get; set; } = new();

    [SettingsSection("Create from existing")]
    public ProvisioningSettings Provisioning { get; set; } = new();

    [SettingsSection("Modernize")]
    public ModernizeSettings Modernize { get; set; } = new();

    [SettingsSection("Attachments")]
    public AttachmentSettings Attachments { get; set; } = new();

    [SettingsSection("Naming")]
    public NamingSettings Naming { get; set; } = new();

    [SettingsSection("Advanced")]
    public AdvancedSettings Advanced { get; set; } = new();
}

public sealed class ConnectionSettings
{
    [Setting("Tenant name",
        "Your Microsoft 365 tenant: the short name (contoso), contoso.onmicrosoft.com, or the tenant GUID. " +
        "Used as the sign-in authority, and to expand site URLs typed as /sites/Name.",
        Required = true)]
    public string TenantName { get; set; } = "";

    [Setting("App (client) ID",
        "The Application (client) ID of YOUR Entra ID app registration, used for interactive sign-in. " +
        "No default is supplied: register an app in your tenant (see README), grant it delegated SharePoint " +
        "permissions with admin consent, and add the redirect URI below under 'Mobile and desktop applications'.",
        Required = true)]
    public string ClientId { get; set; } = "";

    [Setting("Redirect URI",
        "Must match a 'Mobile and desktop applications' redirect URI on the app registration. " +
        "http://localhost works for most registrations; an app already used with PnP.PowerShell -Interactive usually has it.",
        Required = true)]
    public string RedirectUri { get; set; } = "http://localhost";

    [Setting("Sign-in account",
        "Optional. The account to sign in with, e.g. admin@contoso.com. The sign-in page is pre-filled with it, " +
        "and the log warns if a run connects as anyone else. Blank = choose an account at sign-in.")]
    public string LoginHint { get; set; } = "";

    /// <summary>The sign-in details for this tenant.</summary>
    public TenantSignIn ToSignIn() => new(TenantName, ClientId, RedirectUri, LoginHint);
}

/// <summary>
/// The second tenant, used when 'Destination is in another tenant' is ticked on the main window.
/// Only the tenant name is required; App ID and redirect URI fall back to the source tenant's,
/// which suits a multi-tenant app registration consented in both tenants.
/// </summary>
public sealed class DestinationTenantSettings
{
    [Setting("Tenant name",
        "The DESTINATION Microsoft 365 tenant: short name (fabrikam), fabrikam.onmicrosoft.com, or the tenant GUID. " +
        "Required for cross-tenant runs; ignored otherwise.")]
    public string TenantName { get; set; } = "";

    [Setting("App (client) ID",
        "The app registration used to sign in to the destination tenant. Blank = use the source tenant's App ID " +
        "(works for a multi-tenant app that has been consented in both tenants).")]
    public string ClientId { get; set; } = "";

    [Setting("Redirect URI", "Blank = same as the source tenant.")]
    public string RedirectUri { get; set; } = "";

    [Setting("Sign-in account",
        "Optional. The destination tenant account, e.g. admin@fabrikam.com. Pre-fills the second sign-in, and the " +
        "log warns if a run connects to the destination as anyone else.")]
    public string LoginHint { get; set; } = "";

    /// <summary>Destination sign-in details, with blanks filled from the source tenant.</summary>
    public TenantSignIn ToSignIn(ConnectionSettings source) => new(
        TenantName,
        string.IsNullOrWhiteSpace(ClientId) ? source.ClientId : ClientId,
        string.IsNullOrWhiteSpace(RedirectUri) ? source.RedirectUri : RedirectUri,
        LoginHint);
}

/// <summary>Everything one sign-in needs. Built from either tenant section.</summary>
public sealed record TenantSignIn(string TenantName, string ClientId, string RedirectUri, string LoginHint);

public sealed class CrossTenantSettings
{
    [Setting("User mapping CSV",
        "Optional. A CSV with the columns SourceUser,DestinationUser (emails/UPNs) that matches people in the source " +
        "tenant to accounts in the destination tenant. Checked first. 'Generate CSV Template' writes a starter file.",
        IsFile = true)]
    public string UserMappingCsvPath { get; set; } = "";

    [Setting("Domain mappings",
        "One per line, as sourcedomain=destinationdomain (e.g. contoso.com=fabrikam.com). Used for anyone not in the " +
        "mapping CSV: jane@contoso.com becomes jane@fabrikam.com.")]
    public List<string> DomainMappings { get; set; } = new();

    [Setting("Fallback account",
        "Optional. Destination account used when a person cannot be matched (e.g. migration@fabrikam.com). " +
        "Blank = leave the person blank. Every unmatched person is logged to UnresolvedUsers.csv either way.")]
    public string FallbackUser { get; set; } = "";

    [Setting("Upload chunk size (MB)",
        "Files are downloaded from the source tenant and uploaded to the destination. Files larger than this are " +
        "uploaded in chunks of this size.", Min = 1, Max = 100)]
    public int UploadChunkSizeMB { get; set; } = 10;

    [Setting("Keep transfer staging files",
        "Keep the downloaded copies under <run>\\TransferStaging after each upload. Off = each file is deleted once uploaded.")]
    public bool KeepStagingFiles { get; set; } = false;
}

public sealed class PathSettings
{
    [Setting("Output root",
        "Root folder for per-run working output: attachment staging and provisioning schema XML. " +
        "Each run gets a subfolder named <yyyyMMdd_HHmmss>_<Operation>.",
        Required = true, IsFolder = true)]
    public string OutputRoot { get; set; } = @"C:\ContentMigrator";

    [Setting("Log location",
        "Root folder for run logs (OperationLog, ErrorLog, MetadataLog, FolderCreationLog, UnresolvedUsers, " +
        "ProvisioningLog, AttachmentLog, RunMetadata, RunSummary, RunLog.txt). Blank = same as Output root.",
        IsFolder = true)]
    public string LogRoot { get; set; } = "";

    [Setting("Report location",
        "Root folder for reports (inventories, metadata compatibility reports, CSV templates, retry input). " +
        "Blank = same as Output root.",
        IsFolder = true)]
    public string ReportRoot { get; set; } = "";

    [Setting("Write RunLog.txt",
        "Also write every line shown in the on-screen log to RunLog.txt in the run's log folder.")]
    public bool WriteRunLogFile { get; set; } = true;
}

public sealed class CopySettings
{
    [Setting("Overwrite existing (default)", "Default for the 'Overwrite existing' check box on the main window.")]
    public bool OverwriteExisting { get; set; } = false;

    [Setting("Preserve system metadata (default)",
        "Default for the 'Preserve system metadata' check box: replays Created, Modified, Author and Editor.")]
    public bool PreserveSystemMetadata { get; set; } = true;

    [Setting("Create destination if missing (default)",
        "Default for the 'Create destination from source if missing' check box.")]
    public bool CreateDestinationIfMissing { get; set; } = true;

    [Setting("Allow schema mismatch",
        "Cross-site copy/move (copy job): allow the file even when source and destination columns differ.")]
    public bool AllowSchemaMismatch { get; set; } = true;

    [Setting("Allow smaller version limit on destination",
        "Cross-site copy/move (copy job): allow a destination library whose version limit is lower than the source.")]
    public bool AllowSmallerVersionLimitOnDestination { get; set; } = true;

    [Setting("Check destination for duplicates", "Look for the destination file before each copy/move.")]
    public bool CheckDestinationDuplicates { get; set; } = true;

    [Setting("Skip existing when overwrite is off",
        "With duplicate checking on and overwrite off, log existing destination files as Skipped.")]
    public bool SkipExistingWhenOverwriteDisabled { get; set; } = true;

    [Setting("Wait for destination file before metadata",
        "Poll until the copied/moved file is visible before replaying metadata.")]
    public bool WaitForDestinationFileBeforeMetadata { get; set; } = true;

    [Setting("Destination wait: max attempts", "", Min = 1, Max = 1000)]
    public int DestinationFileWaitMaxAttempts { get; set; } = 12;

    [Setting("Destination wait: seconds between attempts", "", Min = 1, Max = 600)]
    public int DestinationFileWaitSeconds { get; set; } = 5;

    [Setting("Copy job timeout (minutes)",
        "Cross-site copy/move uses SharePoint copy jobs. Give up waiting on one file after this many minutes.",
        Min = 1, Max = 720)]
    public int CopyJobTimeoutMinutes { get; set; } = 30;
}

public sealed class ReportSettings
{
    [Setting("Inventory: include folders")]
    public bool IncludeFoldersInInventory { get; set; } = true;

    [Setting("Inventory: include files")]
    public bool IncludeFilesInInventory { get; set; } = true;

    [Setting("Metadata compatibility report",
        "Write MetadataCompatibility_<context>.csv comparing source and destination columns before each operation.")]
    public bool EnableMetadataCompatibilityReport { get; set; } = true;

    [Setting("Log skipped metadata",
        "Write a MetadataLog row for every unsupported column skipped on every list item, and show skipped Modernize columns.")]
    public bool LogMetadataSkipped { get; set; } = true;
}

public sealed class ProvisioningSettings
{
    [Setting("Enable create-from-existing",
        "Allows Operation 9, and creating a missing destination during copy/move/list copy.")]
    public bool EnableCreateDestinationFromSource { get; set; } = true;

    [Setting("Review-only (default)",
        "Default for the 'Provision review-only' check box: write the schema XML and column plan, create nothing.")]
    public bool ReviewOnly { get; set; } = false;

    [Setting("Save provisioning schema XML",
        "Write the source list schema (list settings, columns, views) to ProvisioningSchema_<list>.xml in the run folder.")]
    public bool SaveProvisioningSchemaXml { get; set; } = true;

    [Setting("Column types not provisioned",
        "One per line. These carry cross-site or term-store dependencies and are not recreated.")]
    public List<string> ExcludedFieldTypes { get; set; } = new()
    {
        "Lookup", "LookupMulti", "TaxonomyFieldType", "TaxonomyFieldTypeMulti"
    };

    [Setting("Copy list content types",
        "Bind the source list's content types on the destination when a matching parent content type exists in the destination site.")]
    public bool CopyContentTypes { get; set; } = true;

    [Setting("Replicate views")]
    public bool EnableViewReplication { get; set; } = true;

    [Setting("Replicate personal views", "Off = public views only.")]
    public bool ReplicatePersonalViews { get; set; } = false;

    [Setting("Overwrite existing views", "Off = skip views that already exist on the destination.")]
    public bool OverwriteExistingViews { get; set; } = false;
}

public sealed class ModernizeSettings
{
    [Setting("Enable Modernize List")]
    public bool EnableModernizeList { get; set; } = true;

    [Setting("Copy items after modernizing", "Off = schema and views only, no items.")]
    public bool CopyItems { get; set; } = true;

    [Setting("Enable versioning on the new list")]
    public bool EnableVersioning { get; set; } = true;

    [Setting("Enable attachments on the new list", "Needed for attachment migration.")]
    public bool EnableAttachments { get; set; } = true;

    [Setting("Column group", "Blank = keep the source column group.")]
    public string FieldGroup { get; set; } = "";

    [Setting("Column types not recreated", "One per line.")]
    public List<string> ExcludedFieldTypes { get; set; } = new()
    {
        "Lookup", "LookupMulti", "TaxonomyFieldType", "TaxonomyFieldTypeMulti", "Calculated", "Computed",
        "Attachments", "WorkflowStatus", "WorkflowEventType", "ThreadIndex", "AllDayEvent", "Recurrence",
        "CrossProjectLink"
    };
}

public sealed class AttachmentSettings
{
    [Setting("Copy list item attachments")]
    public bool EnableAttachmentCopy { get; set; } = true;

    [Setting("Keep staging files", "Keep downloaded attachments under <run>\\AttachmentStaging after the run.")]
    public bool KeepStagingFiles { get; set; } = false;

    [Setting("Max attachment size (MB)", "Larger attachments are skipped and logged. 0 = no limit.", Min = 0, Max = 250)]
    public int MaxAttachmentSizeMB { get; set; } = 250;

    [Setting("Attachment failure fails the item",
        "On = an item whose attachment fails is logged as Error in OperationLog.csv.")]
    public bool AttachmentFailureIsItemFailure { get; set; } = false;
}

public sealed class NamingSettings
{
    [Setting("Enforce destination naming",
        "The destination name you type becomes the display title AND the URL (spaces and illegal characters removed).")]
    public bool EnforceDestinationNaming { get; set; } = true;

    [Setting("Illegal URL character pattern", "Regular expression of characters removed from the URL segment.")]
    public string UrlIllegalCharacterPattern { get; set; } = "[~#%&*{}\\\\:<>?/+|\"']";

    [Setting("URL rename failure is fatal",
        "Off = keep the correct title, log a warning about the URL, and continue.")]
    public bool UrlRenameFailureIsFatal { get; set; } = false;
}

public sealed class AdvancedSettings
{
    [Setting("Require site collection admin",
        "Stop before anything is read or changed if the signed-in account is not a Site Collection Admin on a source " +
        "or destination site (in either tenant). The SharePoint Administrator role alone does not make an account a " +
        "site admin, and lists with unique permissions are invisible to non-admins. Off = log a warning and continue.")]
    public bool RequireSiteCollectionAdmin { get; set; } = true;

    [Setting("Page size", "Items requested per page when reading lists.", Min = 100, Max = 5000)]
    public int DefaultPageSize { get; set; } = 2000;

    [Setting("Throttling: max retries", "Retries for HTTP 429/503 responses from SharePoint.", Min = 0, Max = 50)]
    public int MaxRetryAttempts { get; set; } = 10;

    [Setting("Throttling: base delay (seconds)",
        "Used when SharePoint sends no Retry-After header; doubles on each retry.", Min = 1, Max = 300)]
    public int RetryBaseDelaySeconds { get; set; } = 5;

    [Setting("Enable retry-from-log (Operation 7)")]
    public bool EnableFailedItemRetryMode { get; set; } = true;

    [Setting("Enable CSV template generator (Operation 8)")]
    public bool EnableCsvTemplateGenerator { get; set; } = true;

    [Setting("Supported metadata field types", "One per line. Column types whose values are copied.")]
    public List<string> SupportedSimpleFieldTypes { get; set; } = new()
    {
        "Text", "Note", "Choice", "MultiChoice", "Number", "Currency", "DateTime", "Boolean", "URL", "User",
        "UserMulti", "Integer"
    };

    [Setting("System field internal names", "One per line. Never copied as ordinary metadata.")]
    public List<string> SystemFieldInternalNames { get; set; } = new()
    {
        "ID", "GUID", "UniqueId", "ContentType", "ContentTypeId", "FileRef", "FileDirRef", "FileLeafRef",
        "File_x0020_Type", "File_x0020_Size", "FSObjType", "Created", "Modified", "Author", "Editor",
        "CheckoutUser", "CheckedOutTitle", "_CheckinComment", "_UIVersion", "_UIVersionString",
        "_ModerationStatus", "_ModerationComments", "AppAuthor", "AppEditor", "SMTotalSize",
        "SMLastModifiedDate", "owshiddenversion", "Order"
    };
}
