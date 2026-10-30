using System.Text.RegularExpressions;
using Microsoft.SharePoint.Client;
using ContentMigrator.Config;
using ContentMigrator.Infrastructure;
using ContentMigrator.SharePoint;

namespace ContentMigrator.Engine;

/// <summary>
/// Raised when the signed-in account is not a Site Collection Admin on a site the run needs.
/// The window shows its message in a dialog as well as the log.
/// </summary>
public sealed class SiteAdminRequiredException : InvalidOperationException
{
    public SiteAdminRequiredException(string message) : base(message) { }
}

/// <summary>Per-run choices made on the main window (the former GUI check boxes).</summary>
public sealed class RunOptions
{
    public bool OverwriteExisting { get; init; }
    public bool PreserveSystemMetadata { get; init; }
    public bool CreateDestinationIfMissing { get; init; }
    public bool ReviewOnly { get; init; }

    /// <summary>The destination site is in the Destination tenant (Settings > Destination tenant).</summary>
    public bool CrossTenant { get; init; }
}

/// <summary>Loaded list properties (the former Get-PnPList result).</summary>
public sealed class SpListInfo
{
    public List List { get; init; }
    public Guid Id { get; init; }
    public string Title { get; init; }
    public string RootFolderUrl { get; init; }
    public BaseType BaseType { get; init; }
    public int BaseTemplate { get; init; }
    public Guid TemplateFeatureId { get; init; }
    public bool EnableAttachments { get; init; }
    public bool EnableVersioning { get; init; }
    public bool EnableMinorVersions { get; init; }
    public int MajorVersionLimit { get; init; }
    public int MajorWithMinorVersionsLimit { get; init; }
    public bool ContentTypesEnabled { get; init; }
    public string Description { get; init; }
    public bool IsLibrary => BaseType == BaseType.DocumentLibrary;
}

/// <summary>Loaded field properties (the former Get-PnPField result).</summary>
public sealed class SpFieldInfo
{
    public Guid Id { get; init; }
    public string InternalName { get; init; }
    public string Title { get; init; }
    public string TypeAsString { get; init; }
    public bool Hidden { get; init; }
    public bool ReadOnlyField { get; init; }
    public string SchemaXml { get; init; }
}

public sealed class FieldMapEntry
{
    public string InternalName { get; init; }
    public string DisplayName { get; init; }
    public string SourceTypeAsString { get; init; }
    public string DestTypeAsString { get; init; }
}

/// <summary>
/// The migration engine: every SharePoint operation from the v2.3.1 script, rewritten on
/// Microsoft's SharePoint client object model (CSOM). Runs on a background thread; one
/// instance per run.
/// </summary>
public sealed partial class MigrationEngine : IDisposable
{
    private readonly AppSettings _s;
    private readonly RunContext _run;
    private readonly SpoAuthenticator _sourceAuth;
    private readonly SpoAuthenticator _destinationAuth;
    private readonly RunOptions _opt;
    private readonly UserMapper _userMapper;
    private readonly CancellationToken _ct;
    private readonly Dictionary<string, SpConnection> _connections = new(StringComparer.OrdinalIgnoreCase);
    private readonly HashSet<string> _supportedTypes;
    private readonly HashSet<string> _systemFields;
    private readonly Regex _urlIllegal;

    /// <param name="destinationAuth">The destination tenant's sign-in for a cross-tenant run; ignored otherwise.</param>
    public MigrationEngine(AppSettings settings, RunContext run, SpoAuthenticator sourceAuth, SpoAuthenticator destinationAuth,
        RunOptions options, CancellationToken cancellationToken)
    {
        _s = settings;
        _run = run;
        _sourceAuth = sourceAuth;
        _opt = options;
        _destinationAuth = options.CrossTenant
            ? destinationAuth ?? throw new ArgumentNullException(nameof(destinationAuth), "A cross-tenant run needs a destination sign-in.")
            : sourceAuth;
        _userMapper = options.CrossTenant ? new UserMapper(settings.CrossTenant, run) : null;
        _ct = cancellationToken;
        _supportedTypes = new HashSet<string>(settings.Advanced.SupportedSimpleFieldTypes, StringComparer.OrdinalIgnoreCase);
        _systemFields = new HashSet<string>(settings.Advanced.SystemFieldInternalNames, StringComparer.OrdinalIgnoreCase);
        _urlIllegal = new Regex(settings.Naming.UrlIllegalCharacterPattern ?? "");
    }

    private bool CancelRequested => _ct.IsCancellationRequested;

    private void Log(string message, MigratorLogLevel level = MigratorLogLevel.Info) => _run.Log(message, level);

    public void Dispose()
    {
        foreach (var c in _connections.Values) c.Dispose();
        _connections.Clear();
    }

    // ======================================================================= preflight

    /// <summary>Configuration sanity before anything connects.</summary>
    public void CheckPrerequisites()
    {
        Log("Running pre-flight checks...");
        var problems = SettingsStore.Validate(_s);
        if (_opt.CrossTenant) problems.AddRange(SettingsStore.ValidateCrossTenant(_s));
        if (problems.Count > 0)
        {
            throw new InvalidOperationException("Settings need attention before running: " + string.Join(" ", problems));
        }
        Log($"Runtime: {_run.RuntimeVersion}", MigratorLogLevel.Success);
        Log($"SharePoint CSOM: {_run.CsomVersion}", MigratorLogLevel.Success);
        Log($"Source tenant: {_s.Connection.TenantName} | App ID: {_s.Connection.ClientId}", MigratorLogLevel.Success);
        if (_opt.CrossTenant)
        {
            var d = _s.Destination.ToSignIn(_s.Connection);
            Log($"Destination tenant: {d.TenantName} | App ID: {d.ClientId}", MigratorLogLevel.Success);
            Log("CROSS-TENANT run: files are downloaded from the source tenant and uploaded to the destination " +
                "(latest version only). Move is not available across tenants.", MigratorLogLevel.Info);
            _userMapper.Load();
        }
        if (_s.Provisioning.EnableCreateDestinationFromSource)
            Log("Create-from-existing is available (native CSOM schema rebuild).", MigratorLogLevel.Success);
        if (_s.Modernize.EnableModernizeList)
            Log("Modernize List is available.", MigratorLogLevel.Success);
        if (_s.Attachments.EnableAttachmentCopy)
            Log("List item attachment copy is available.", MigratorLogLevel.Success);
    }

    // ======================================================================= connections

    /// <summary>
    /// Connects to a site in the given tenant. Connections are reused for the same site within a
    /// run. Same-tenant runs treat both sides as the source tenant, so one sign-in serves both.
    /// </summary>
    public SpConnection Connect(string siteUrl, TenantSide side = TenantSide.Source)
    {
        var effectiveSide = _opt.CrossTenant ? side : TenantSide.Source;
        var isDestination = effectiveSide == TenantSide.Destination;
        var tenantName = isDestination ? _s.Destination.TenantName : _s.Connection.TenantName;
        var auth = isDestination ? _destinationAuth : _sourceAuth;
        var expected = ((isDestination ? _s.Destination.LoginHint : _s.Connection.LoginHint) ?? "").Trim();
        var tenantLabel = _opt.CrossTenant ? $" ({effectiveSide.ToString().ToLowerInvariant()} tenant)" : "";

        siteUrl = SpoAuthenticator.ResolveSiteUrl(siteUrl, tenantName);
        var key = effectiveSide + "|" + siteUrl;
        if (_connections.TryGetValue(key, out var existing)) return existing;

        Log($"Connecting to {siteUrl}{tenantLabel}");
        try
        {
            var ctx = auth.CreateContext(siteUrl);
            var conn = new SpConnection(siteUrl, ctx, _s.Advanced.MaxRetryAttempts, _s.Advanced.RetryBaseDelaySeconds,
                effectiveSide, auth);
            conn.Open();
            var who = conn.CurrentUserName;
            Log($"Connected to {siteUrl}{tenantLabel} as {who}", MigratorLogLevel.Success);
            if (expected.Length > 0 && !string.Equals(who, expected, StringComparison.OrdinalIgnoreCase))
            {
                var section = isDestination ? "Destination tenant" : "Source tenant";
                Log($"Signed in as {who}, but Settings > {section} > Sign-in account is {expected}. " +
                    "Restart the app and choose the configured account if this is not intended.", MigratorLogLevel.Warn);
            }
            if (!string.IsNullOrWhiteSpace(who))
            {
                if (isDestination && string.IsNullOrEmpty(_run.DestinationOperatorUpn)) _run.DestinationOperatorUpn = who;
                if (!isDestination && _run.OperatorUpn == "Unknown") _run.OperatorUpn = who;
            }
            CheckSiteAdmin(conn);
            _connections[key] = conn;
            return conn;
        }
        catch (SiteAdminRequiredException)
        {
            throw; // already logged by CheckSiteAdmin
        }
        catch (Exception ex)
        {
            _run.AddErrorLog("Connection", "Connect", siteUrl, "", ex);
            throw;
        }
    }

    /// <summary>Stops a Move whose destination is in another tenant (there is no cross-tenant move).</summary>
    private bool BlockCrossTenantMove(string operationType, string modeName)
    {
        if (!_opt.CrossTenant || operationType != "Move") return false;
        const string message = "Move is not available across tenants: SharePoint has no cross-tenant move, and this " +
                               "tool never deletes source content in another tenant. Run the Copy operation instead.";
        Log(message, MigratorLogLevel.Error);
        _run.AddErrorLog(modeName, "CrossTenantMoveBlocked", "", "", message, "CrossTenantMoveBlocked");
        return true;
    }

    /// <summary>
    /// The script's prerequisite ("You MUST be a Site Collection Admin on BOTH the SOURCE and
    /// DESTINATION sites"), now enforced. The SharePoint Administrator role does not make an
    /// account a site admin, and lists with unique permissions are invisible to non-admins,
    /// so without this check the run fails later with a misleading "not found".
    /// </summary>
    private void CheckSiteAdmin(SpConnection conn)
    {
        if (conn.CurrentUserIsSiteAdmin)
        {
            Log($"{conn.CurrentUserName} is a Site Collection Admin on {conn.SiteUrl}.", MigratorLogLevel.Success);
            return;
        }

        var who = conn.CurrentUserName;
        if (!_s.Advanced.RequireSiteCollectionAdmin)
        {
            Log($"{who} is NOT a Site Collection Admin on {conn.SiteUrl}. Continuing because 'Require site collection admin' " +
                "is off; lists with unique permissions may be reported as not found and some copies may fail.", MigratorLogLevel.Warn);
            return;
        }

        var message =
            $"The signed-in account {who} is not a Site Collection Admin on {conn.SiteUrl}.{Environment.NewLine}{Environment.NewLine}" +
            "The SharePoint Administrator role does not make an account a site admin of each site. Add the account as a " +
            "site admin (SharePoint admin center > Active sites > the site > Membership > Site admins), or run:" +
            $"{Environment.NewLine}{Environment.NewLine}" +
            $"Set-SPOUser -Site {conn.SiteUrl} -LoginName {who} -IsSiteCollectionAdmin $true" +
            $"{Environment.NewLine}{Environment.NewLine}" +
            "Then run the operation again. Nothing was read, copied, moved or changed.";
        Log($"{who} is not a Site Collection Admin on {conn.SiteUrl}. Operation stopped before any data was touched.", MigratorLogLevel.Error);
        _run.AddErrorLog("Connection", "SiteAdminCheck", conn.SiteUrl, "",
            $"{who} is not a Site Collection Admin on {conn.SiteUrl}.", "SiteAdminRequired");
        throw new SiteAdminRequiredException(message);
    }

    // ======================================================================= list helpers

    /// <summary>One row of the site's list catalogue (cheap properties only).</summary>
    private sealed class ListCatalogEntry
    {
        public Guid Id { get; init; }
        public string Title { get; init; }
        public string RootFolderUrl { get; init; }
        public bool Hidden { get; init; }
    }

    /// <summary>
    /// Get-PnPList -Identity. Reads the site's lists once (Id, Title, RootFolder) and matches the
    /// operator's input against them, so "not found" never depends on catching an exception.
    /// Accepts: display title, URL name ("Shared Documents", "Lists/Tasks"), server-relative or
    /// absolute URL (view pages like /Forms/AllItems.aspx are trimmed), or the list GUID.
    /// Returns null only when no list matches. Any real error (permissions, network) is thrown.
    /// </summary>
    private SpListInfo TryGetList(SpConnection c, string identity) => TryGetList(c, identity, out _);

    private SpListInfo TryGetList(SpConnection c, string identity, out List<ListCatalogEntry> catalog)
    {
        catalog = null;
        if (string.IsNullOrWhiteSpace(identity)) return null;
        identity = identity.Trim();

        var lists = c.Web.Lists;
        c.Context.Load(lists, ls => ls.Include(l => l.Id, l => l.Title, l => l.Hidden, l => l.RootFolder.ServerRelativeUrl));
        c.ExecuteQuery();
        catalog = lists.Select(l => new ListCatalogEntry
        {
            Id = l.Id,
            Title = l.Title,
            Hidden = l.Hidden,
            RootFolderUrl = NormalizeServerRelativeUrl(l.RootFolder.ServerRelativeUrl)
        }).ToList();

        var match = MatchList(catalog, identity, c);
        return match == null ? null : LoadListDetails(c, match);
    }

    private ListCatalogEntry MatchList(List<ListCatalogEntry> catalog, string identity, SpConnection c)
    {
        if (Guid.TryParse(identity.Trim('{', '}'), out var id))
        {
            var byId = catalog.FirstOrDefault(l => l.Id == id);
            if (byId != null) return byId;
        }

        // 1. Display title (exact, then case-insensitive).
        var byTitle = catalog.FirstOrDefault(l => string.Equals(l.Title, identity, StringComparison.Ordinal))
                      ?? catalog.FirstOrDefault(l => string.Equals(l.Title, identity, StringComparison.OrdinalIgnoreCase));
        if (byTitle != null) return byTitle;

        // 2. Anything URL-shaped: reduce to a server-relative path.
        var path = identity;
        if (Uri.TryCreate(path, UriKind.Absolute, out var abs) && (abs.Scheme == "https" || abs.Scheme == "http"))
        {
            path = abs.AbsolutePath;
        }
        path = Uri.UnescapeDataString(path).Replace('\\', '/').Trim();
        path = Regex.Replace(path, @"/Forms(/[^/]*)?$", "", RegexOptions.IgnoreCase);      // library view page
        path = Regex.Replace(path, @"/[^/]+\.aspx$", "", RegexOptions.IgnoreCase);         // list view page
        path = path.TrimEnd('/');

        if (path.StartsWith("/"))
        {
            var norm = NormalizeServerRelativeUrl(path);
            var byPath = catalog.FirstOrDefault(l => string.Equals(l.RootFolderUrl, norm, StringComparison.OrdinalIgnoreCase));
            if (byPath != null) return byPath;
        }

        // 3. Site-relative URL ("Lists/Tasks", "Shared Documents") or just the URL leaf.
        var siteRelative = path.TrimStart('/');
        var webPrefix = (c.WebServerRelativeUrl ?? "").Trim('/');
        if (webPrefix.Length > 0 && siteRelative.StartsWith(webPrefix + "/", StringComparison.OrdinalIgnoreCase))
        {
            siteRelative = siteRelative.Substring(webPrefix.Length + 1);
        }
        foreach (var candidate in new[] { siteRelative, "Lists/" + siteRelative })
        {
            var full = NormalizeServerRelativeUrl((c.WebServerRelativeUrl ?? "") + "/" + candidate);
            var hit = catalog.FirstOrDefault(l => string.Equals(l.RootFolderUrl, full, StringComparison.OrdinalIgnoreCase));
            if (hit != null) return hit;
        }

        var leaf = GetLeafName(siteRelative);
        var byLeaf = catalog.Where(l => string.Equals(GetLeafName(l.RootFolderUrl), leaf, StringComparison.OrdinalIgnoreCase)).ToList();
        return byLeaf.Count == 1 ? byLeaf[0] : null;
    }

    /// <summary>
    /// Loads the properties the engine uses. Core properties are one request; the optional ones
    /// are a second request whose failure is logged but never hides the list.
    /// </summary>
    private SpListInfo LoadListDetails(SpConnection c, ListCatalogEntry entry)
    {
        var list = c.Web.Lists.GetById(entry.Id);
        c.Context.Load(list, l => l.Id, l => l.Title, l => l.BaseType, l => l.BaseTemplate, l => l.Description);
        c.ExecuteQuery();

        Guid templateFeatureId = Guid.Empty;
        bool enableAttachments = false, enableVersioning = false, enableMinor = false, contentTypesEnabled = false;
        int majorLimit = 0, majorWithMinorLimit = 0;

        var isLibrary = list.BaseType == BaseType.DocumentLibrary;
        var optional = new List<(string Name, System.Linq.Expressions.Expression<Func<List, object>> Property, Action Read)>
        {
            ("EnableVersioning", l => l.EnableVersioning, () => enableVersioning = list.EnableVersioning),
            ("MajorVersionLimit", l => l.MajorVersionLimit, () => majorLimit = list.MajorVersionLimit),
            ("ContentTypesEnabled", l => l.ContentTypesEnabled, () => contentTypesEnabled = list.ContentTypesEnabled),
            ("TemplateFeatureId", l => l.TemplateFeatureId, () => templateFeatureId = list.TemplateFeatureId)
        };
        if (isLibrary)
        {
            optional.Add(("EnableMinorVersions", l => l.EnableMinorVersions, () => enableMinor = list.EnableMinorVersions));
            optional.Add(("MajorWithMinorVersionsLimit", l => l.MajorWithMinorVersionsLimit,
                () => majorWithMinorLimit = list.MajorWithMinorVersionsLimit));
        }
        else
        {
            optional.Add(("EnableAttachments", l => l.EnableAttachments, () => enableAttachments = list.EnableAttachments));
        }

        // One round trip normally; property by property only if the batch is refused.
        try
        {
            c.Context.Load(list, optional.Select(o => o.Property).ToArray());
            c.ExecuteQuery();
            foreach (var o in optional) o.Read();
        }
        catch (ServerException batchError)
        {
            Log($"List '{entry.Title}': reading list settings in one request failed ({batchError.Message}). Retrying one at a time.",
                MigratorLogLevel.Debug);
            foreach (var o in optional)
            {
                try
                {
                    c.Context.Load(list, o.Property);
                    c.ExecuteQuery();
                    o.Read();
                }
                catch (ServerException ex)
                {
                    Log($"List '{entry.Title}': could not read {o.Name} ({ex.Message}). Using the default.", MigratorLogLevel.Warn);
                }
            }
        }

        return new SpListInfo
        {
            List = list,
            Id = list.Id,
            Title = list.Title,
            RootFolderUrl = entry.RootFolderUrl,
            BaseType = list.BaseType,
            BaseTemplate = list.BaseTemplate,
            TemplateFeatureId = templateFeatureId,
            EnableAttachments = enableAttachments,
            EnableVersioning = enableVersioning,
            EnableMinorVersions = enableMinor,
            MajorVersionLimit = majorLimit,
            MajorWithMinorVersionsLimit = majorWithMinorLimit,
            ContentTypesEnabled = contentTypesEnabled,
            Description = list.Description
        };
    }

    private SpListInfo GetListOrThrow(SpConnection c, string identity)
    {
        var info = TryGetList(c, identity, out var catalog);
        if (info != null) return info;

        var visible = (catalog ?? new List<ListCatalogEntry>())
            .Where(l => !l.Hidden)
            .Select(l => $"'{l.Title}' ({GetLeafName(l.RootFolderUrl)})")
            .OrderBy(s => s, StringComparer.OrdinalIgnoreCase)
            .ToList();
        var who = c.CurrentUserName;
        var hint = visible.Count == 0
            ? " The site returned no visible lists."
            : " Lists and libraries visible to this account: " + string.Join(", ", visible) + ".";
        throw new InvalidOperationException(
            $"List or library '{identity}' was not found on {c.SiteUrl} for the signed-in account {who}.{hint} " +
            "If it does exist, this account cannot see it: check that you signed in with the right account " +
            "(restart the app to pick a different one) and that the account has access to that list, which may have unique permissions.");
    }

    private bool ListExists(SpConnection c, string identity) => TryGetList(c, identity) != null;

    private List<SpFieldInfo> GetFields(SpConnection c, List list)
    {
        var fields = list.Fields;
        c.Context.Load(fields, fs => fs.Include(
            f => f.Id, f => f.InternalName, f => f.Title, f => f.TypeAsString, f => f.Hidden,
            f => f.ReadOnlyField, f => f.SchemaXml));
        c.ExecuteQuery();
        return fields.Select(f => new SpFieldInfo
        {
            Id = f.Id,
            InternalName = f.InternalName,
            Title = f.Title,
            TypeAsString = f.TypeAsString,
            Hidden = f.Hidden,
            ReadOnlyField = f.ReadOnlyField,
            SchemaXml = f.SchemaXml
        }).ToList();
    }

    /// <summary>
    /// Get-PnPListItem -PageSize: every item (folders included, as the script did), paged and
    /// ordered by ID so large lists stay under the list view threshold.
    /// </summary>
    private List<ListItem> GetAllItems(SpConnection c, List list, IEnumerable<string> viewFields = null)
    {
        var result = new List<ListItem>();
        var vf = viewFields == null
            ? ""
            : "<ViewFields>" + string.Concat(viewFields.Select(f => $"<FieldRef Name='{f}'/>")) + "</ViewFields>";
        ListItemCollectionPosition position = null;
        do
        {
            var query = new CamlQuery
            {
                ViewXml = "<View Scope='RecursiveAll'><Query><OrderBy><FieldRef Name='ID' Ascending='TRUE'/></OrderBy></Query>" +
                          vf + $"<RowLimit Paged='TRUE'>{_s.Advanced.DefaultPageSize}</RowLimit></View>",
                ListItemCollectionPosition = position
            };
            var items = list.GetItems(query);
            c.Context.Load(items);
            c.ExecuteQuery();
            result.AddRange(items);
            position = items.ListItemCollectionPosition;
            if (CancelRequested) break;
        } while (position != null);
        return result;
    }

    // ======================================================================= path helpers

    public static string NormalizeServerRelativeUrl(string url)
    {
        if (string.IsNullOrWhiteSpace(url)) return null;
        var n = url.Trim().Replace('\\', '/');
        n = Regex.Replace(n, "/+", "/");
        if (!n.StartsWith("/")) n = "/" + n;
        return n.TrimEnd('/');
    }

    public static string JoinUrlPath(string left, string right)
    {
        if (string.IsNullOrWhiteSpace(right)) return NormalizeServerRelativeUrl(left);
        var l = (NormalizeServerRelativeUrl(left) ?? "").TrimEnd('/');
        var r = right.Trim().Trim('/');
        return $"{l}/{r}";
    }

    public static string GetRelativePathUnderRoot(string itemServerRelativeUrl, string libraryRootServerRelativeUrl)
    {
        var item = NormalizeServerRelativeUrl(itemServerRelativeUrl) ?? "";
        var root = NormalizeServerRelativeUrl(libraryRootServerRelativeUrl) ?? "";
        if (string.Equals(item, root, StringComparison.OrdinalIgnoreCase)) return "";
        if (item.StartsWith(root + "/", StringComparison.OrdinalIgnoreCase)) return item.Substring(root.Length).TrimStart('/');
        return item.TrimStart('/');
    }

    public static string GetParentFolderRelativePath(string relativeFilePath)
    {
        if (string.IsNullOrWhiteSpace(relativeFilePath)) return "";
        var clean = relativeFilePath.Trim('/');
        var lastSlash = clean.LastIndexOf('/');
        return lastSlash < 0 ? "" : clean.Substring(0, lastSlash);
    }

    public static string GetLeafName(string path)
    {
        if (string.IsNullOrWhiteSpace(path)) return "";
        return path.TrimEnd('/').Split('/').Last();
    }

    private static bool SameSite(string a, string b) =>
        string.Equals((a ?? "").TrimEnd('/'), (b ?? "").TrimEnd('/'), StringComparison.OrdinalIgnoreCase);

    // ======================================================================= naming + guard (v2.3.1)

    /// <summary>
    /// "Accounting Reconciliation Forms 2026" -> "AccountingReconciliationForms2026".
    /// Removes all whitespace and illegal characters, trims dots, keeps casing, caps at 100.
    /// </summary>
    public string GetUrlSafeName(string title)
    {
        if (string.IsNullOrWhiteSpace(title)) return "";
        var name = Regex.Replace(title.Trim(), @"\s", "");
        name = _urlIllegal.Replace(name, "");
        name = name.Trim('.');
        if (name.Length > 100) name = name.Substring(0, 100);
        return name;
    }

    /// <summary>
    /// Same source / same destination guard. Always on (not configurable).
    /// Returns true when the operation is SAFE to continue.
    /// </summary>
    private bool CheckSourceAndDestinationDiffer(string sourceSiteUrl, string sourceList, string destinationSiteUrl,
        string destinationList, string modeName)
    {
        var srcSite = (sourceSiteUrl ?? "").TrimEnd('/');
        var dstSite = (destinationSiteUrl ?? "").TrimEnd('/');
        if (!string.Equals(srcSite, dstSite, StringComparison.OrdinalIgnoreCase)) return true;

        var srcName = (sourceList ?? "").Trim();
        var dstName = (destinationList ?? "").Trim();
        var srcUrl = GetUrlSafeName(srcName);
        var dstUrl = GetUrlSafeName(dstName);

        if (string.Equals(srcName, dstName, StringComparison.OrdinalIgnoreCase) ||
            (!string.IsNullOrWhiteSpace(srcUrl) && string.Equals(srcUrl, dstUrl, StringComparison.OrdinalIgnoreCase)))
        {
            Log("SOURCE AND DESTINATION ARE THE SAME. Operation stopped before any data was touched.", MigratorLogLevel.Error);
            Log($"  Site: {srcSite}", MigratorLogLevel.Error);
            Log($"  Source list/library:      '{srcName}'  (url: {srcUrl})", MigratorLogLevel.Error);
            Log($"  Destination list/library: '{dstName}'  (url: {dstUrl})", MigratorLogLevel.Error);
            Log("  The destination site URL was blank or matched the source, so this is a same-site operation. " +
                "Supply a DIFFERENT destination list/library name, or a different destination site.", MigratorLogLevel.Error);

            _run.AddErrorLog(modeName, "SameSourceDestinationGuard", $"{srcSite} :: {srcName}", $"{dstSite} :: {dstName}",
                "Source and destination resolve to the same list/library on the same site. Operation aborted.",
                "SameSourceDestination");
            _run.AddProvisionLog(srcSite, srcName, dstSite, dstName, "Guard", "", "SameSourceDestination", "Blocked", "",
                "Source and destination are identical on the same site. Nothing was copied, moved, or created.");
            return false;
        }
        return true;
    }

    /// <summary>A one-line preflight summary in the log.</summary>
    private void LogPreflight(string modeName, string operationType, string sourceSiteUrl, string sourceName,
        string destinationSiteUrl, string destinationName, int itemCount, string notes = "")
    {
        Log($"Preflight: {modeName} / {operationType} | Source: {sourceSiteUrl} :: {sourceName} | " +
            $"Destination: {destinationSiteUrl} :: {destinationName} | Items: {itemCount} | Overwrite: {_opt.OverwriteExisting}");
        if (!string.IsNullOrWhiteSpace(notes)) Log($"Preflight notes: {notes}", MigratorLogLevel.Debug);
    }
}
