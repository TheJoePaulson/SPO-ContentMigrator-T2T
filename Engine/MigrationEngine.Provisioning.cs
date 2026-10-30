using System.Text.RegularExpressions;
using System.Xml;
using Microsoft.SharePoint.Client;
using ContentMigrator.Infrastructure;
using ContentMigrator.SharePoint;

namespace ContentMigrator.Engine;

public sealed partial class MigrationEngine
{
    private const string NativeMethod = "NativeCsom";

    // ======================================================================= destination gate

    /// <summary>
    /// Returns the destination list, creating it from the
    /// source schema when it is missing and 'Create destination from source if missing' is on.
    /// Returns null when the operation cannot proceed.
    /// </summary>
    private SpListInfo ConfirmDestinationOrProvision(SpConnection src, string sourceList, SpConnection dst, string destinationList)
    {
        var existing = TryGetList(dst, destinationList);
        if (existing != null) return existing;

        Log($"Destination list/library '{destinationList}' was not found on the destination web.", MigratorLogLevel.Warn);
        if (!_s.Provisioning.EnableCreateDestinationFromSource)
        {
            Log("Create-from-existing is disabled in Settings. Cannot continue without an existing destination.", MigratorLogLevel.Error);
            return null;
        }
        if (!_opt.CreateDestinationIfMissing)
        {
            Log("Destination is missing and 'Create destination from source if missing' is unchecked. Aborting this operation.", MigratorLogLevel.Error);
            return null;
        }
        Log("Destination missing; creating it from the source schema.");

        var (status, created) = EnsureDestinationListFromSource(src, sourceList, dst, destinationList);
        switch (status)
        {
            case "Created":
            case "AlreadyExists":
                return created ?? TryGetList(dst, destinationList);
            case "ReviewOnly":
                Log("Provisioning ran in review-only mode. Destination was not created, so the copy/move cannot proceed this run.", MigratorLogLevel.Warn);
                return null;
            default:
                Log($"Provisioning did not complete successfully (status: {status}). Aborting operation.", MigratorLogLevel.Error);
                return null;
        }
    }

    /// <summary>GUI "CreateFromExisting" (Operation 9).</summary>
    public void RunCreateFromExisting(string sourceSiteUrl, string sourceList, string destinationSiteUrl, string destinationList)
    {
        var src = Connect(sourceSiteUrl);
        var dst = Connect(destinationSiteUrl, TenantSide.Destination);
        var (status, created) = EnsureDestinationListFromSource(src, sourceList, dst, destinationList);
        Log($"Create-from-existing completed with status: {status}");
        if (status is "Created" or "AlreadyExists")
        {
            var dstList = created ?? TryGetList(dst, destinationList);
            var srcList = TryGetList(src, sourceList);
            if (dstList != null && srcList != null)
            {
                WriteFieldCompatibilityReport(src, srcList, dst, dstList, "CreateFromExisting");
            }
        }
    }

    // ======================================================================= create from existing

    /// <summary>
    /// Creates the destination list from the source schema:
    ///   1. Create the list with the SOURCE base template, title = typed name, URL = URL-safe name.
    ///   2. Copy versioning / attachment / content-type settings.
    ///   3. Add the source's custom columns from their SchemaXml (InternalName preserved).
    ///      Lookup and Taxonomy columns are skipped, as the v2.2.5 template strip did.
    ///   4. Bind the source's list content types where the parent type exists on the destination.
    ///   5. Rebuild the views natively.
    /// Returns Created | AlreadyExists | ReviewOnly | Skipped | Error, plus the destination list.
    /// </summary>
    private (string Status, SpListInfo Destination) EnsureDestinationListFromSource(SpConnection src, string sourceList,
        SpConnection dst, string destinationList)
    {
        string handlers = "Fields,ContentTypes,Views";

        if (!_s.Provisioning.EnableCreateDestinationFromSource)
        {
            Log("Create-from-existing is disabled in Settings. Skipping provisioning.", MigratorLogLevel.Warn);
            return ("Skipped", null);
        }
        if (!CheckSourceAndDestinationDiffer(src.SiteUrl, sourceList, dst.SiteUrl, destinationList, "Provisioning"))
        {
            return ("Error", null);
        }

        var existing = TryGetList(dst, destinationList);
        if (existing != null)
        {
            Log($"Destination list/library '{destinationList}' already exists. Provisioning skipped (idempotent).");
            _run.AddProvisionLog(src.SiteUrl, sourceList, dst.SiteUrl, destinationList, NativeMethod, handlers,
                "Provision", "AlreadyExists", "", "Destination already present; left untouched.");
            return ("AlreadyExists", existing);
        }

        var srcList = TryGetList(src, sourceList);
        if (srcList == null)
        {
            Log($"Source list/library '{sourceList}' was not found. Cannot provision destination.", MigratorLogLevel.Error);
            _run.AddProvisionLog(src.SiteUrl, sourceList, dst.SiteUrl, destinationList, NativeMethod, handlers,
                "Provision", "Error", "", "Source list not found.");
            return ("Error", null);
        }

        // ---- plan the columns
        var sourceFields = GetFields(src, srcList.List);
        var excluded = new HashSet<string>(_s.Provisioning.ExcludedFieldTypes, StringComparer.OrdinalIgnoreCase);
        var plan = PlanColumns(sourceFields, excluded, "provisioning");
        // Calculated columns reference other columns, so add them last.
        var toCreate = plan.Where(p => p.Include).OrderBy(p => p.Field.TypeAsString == "Calculated" ? 1 : 0).ToList();
        var toSkip = plan.Where(p => !p.Include).ToList();

        // ---- schema snapshot for review / audit
        var schemaPath = "";
        if (_s.Provisioning.SaveProvisioningSchemaXml || _opt.ReviewOnly)
        {
            schemaPath = WriteSchemaSnapshot(src, srcList, sourceFields);
            Log($"Source schema saved: {schemaPath}", MigratorLogLevel.Success);
        }

        foreach (var skip in toSkip.Where(s => s.LogSkip))
        {
            Log($"Skipped column for provisioning: {skip.Field.InternalName} ({skip.Field.TypeAsString}) - {skip.Reason}", MigratorLogLevel.Warn);
            _run.AddProvisionLog(src.SiteUrl, sourceList, dst.SiteUrl, destinationList, NativeMethod, handlers,
                "SkipColumn", "Skipped", schemaPath,
                $"Column '{skip.Field.InternalName}' ({skip.Field.TypeAsString}) skipped: {skip.Reason}");
        }

        if (_opt.ReviewOnly)
        {
            foreach (var c in toCreate)
            {
                _run.AddProvisionLog(src.SiteUrl, sourceList, dst.SiteUrl, destinationList, NativeMethod, handlers,
                    "PlanColumn", "ReviewOnly", schemaPath, $"Would create column '{c.Field.InternalName}' ({c.Field.TypeAsString}).");
            }
            Log("Review-only mode is enabled. Schema saved but destination NOT created.", MigratorLogLevel.Warn);
            Log("Review the XML, then run again with 'Provision review-only' unchecked to apply.", MigratorLogLevel.Warn);
            _run.AddProvisionLog(src.SiteUrl, sourceList, dst.SiteUrl, destinationList, NativeMethod, handlers,
                "ExtractOnly", "ReviewOnly", schemaPath, "Review-only gate active; destination not created.");
            return ("ReviewOnly", null);
        }

        // ---- 1. create the list
        Log($"Creating '{destinationList}' from the '{srcList.Title}' schema (base template {srcList.BaseTemplate})...");
        var warnings = 0;
        SpListInfo dstList;
        try
        {
            dstList = CreateList(dst, destinationList, srcList.BaseTemplate, srcList.TemplateFeatureId, srcList.IsLibrary,
                srcList.Description);
        }
        catch (Exception ex)
        {
            // A base template that is not available on the destination (feature not activated,
            // retired template) falls back to the generic list or library.
            var fallback = srcList.IsLibrary ? (int)ListTemplateType.DocumentLibrary : (int)ListTemplateType.GenericList;
            if (srcList.BaseTemplate == fallback) return ProvisionFailed(src, sourceList, dst, destinationList, schemaPath, ex);

            Log($"Base template {srcList.BaseTemplate} could not be used ({ex.Message}). Falling back to template {fallback}.", MigratorLogLevel.Warn);
            _run.AddProvisionLog(src.SiteUrl, sourceList, dst.SiteUrl, destinationList, NativeMethod, handlers,
                "CreateList", "TemplateFallback", schemaPath, $"Template {srcList.BaseTemplate} failed: {ex.Message}. Used {fallback}.");
            warnings++;
            try
            {
                dstList = CreateList(dst, destinationList, fallback, Guid.Empty, srcList.IsLibrary, srcList.Description);
            }
            catch (Exception ex2)
            {
                return ProvisionFailed(src, sourceList, dst, destinationList, schemaPath, ex2);
            }
        }

        try
        {
            // ---- 2. list settings
            try
            {
                var l = dstList.List;
                l.EnableVersioning = srcList.EnableVersioning;
                if (srcList.EnableVersioning && srcList.MajorVersionLimit > 0) l.MajorVersionLimit = srcList.MajorVersionLimit;
                if (srcList.IsLibrary)
                {
                    l.EnableMinorVersions = srcList.EnableMinorVersions;
                    if (srcList.EnableMinorVersions && srcList.MajorWithMinorVersionsLimit > 0)
                        l.MajorWithMinorVersionsLimit = srcList.MajorWithMinorVersionsLimit;
                }
                else
                {
                    l.EnableAttachments = srcList.EnableAttachments;
                }
                l.ContentTypesEnabled = srcList.ContentTypesEnabled;
                l.Update();
                dst.ExecuteQuery();
            }
            catch (Exception ex)
            {
                warnings++;
                Log($"Could not copy list settings (versioning/attachments/content types): {ex.Message}", MigratorLogLevel.Warn);
            }

            // ---- 3. columns
            var (created, already, failed) = AddColumns(dst, dstList, toCreate, src.SiteUrl, sourceList, destinationList,
                NativeMethod, fieldGroup: null);
            warnings += failed;

            // ---- 4. content types
            if (_s.Provisioning.CopyContentTypes && srcList.ContentTypesEnabled)
            {
                warnings += CopyListContentTypes(src, srcList, dst, dstList, destinationList);
            }

            // ---- 5. naming (idempotent) and views
            var finalName = SetListTitleAndUrl(dst, dstList.Title, destinationList, "Provisioning");
            var finalList = TryGetList(dst, finalName) ?? dstList;
            CopyListViews(src, srcList, dst, finalList);

            var status = warnings > 0 ? "CreatedWithWarnings" : "Created";
            _run.AddProvisionLog(src.SiteUrl, sourceList, dst.SiteUrl, destinationList, NativeMethod, handlers,
                "Provision", status, schemaPath,
                $"Destination created from source schema. ColumnsCreated={created}; AlreadyExisted={already}; Failed={failed}; Skipped={toSkip.Count}.");
            Log($"Destination '{destinationList}' created from source schema.", MigratorLogLevel.Success);
            return ("Created", finalList);
        }
        catch (Exception ex)
        {
            // The list exists even though a later step threw: keep it usable (v2.2.5 behaviour).
            Log($"Provisioning reported an error, but destination list '{dstList.Title}' EXISTS.", MigratorLogLevel.Warn);
            Log($"Error was: {ex.Message}", MigratorLogLevel.Warn);
            Log("Views or some columns may be incomplete - verify them in SharePoint. Continuing.", MigratorLogLevel.Warn);
            _run.AddProvisionLog(src.SiteUrl, sourceList, dst.SiteUrl, destinationList, NativeMethod, handlers,
                "Provision", "CreatedWithWarnings", schemaPath, $"A step failed but the destination list exists. {ex.Message}");
            _run.AddErrorLog("Provisioning", "EnsureDestinationListFromSource", $"{src.SiteUrl} :: {sourceList}",
                $"{dst.SiteUrl} :: {destinationList}", "Non-fatal: " + ex.Message, ex.GetType().FullName);
            return ("Created", TryGetList(dst, destinationList) ?? dstList);
        }
    }

    private (string, SpListInfo) ProvisionFailed(SpConnection src, string sourceList, SpConnection dst,
        string destinationList, string schemaPath, Exception ex)
    {
        _run.AddProvisionLog(src.SiteUrl, sourceList, dst.SiteUrl, destinationList, NativeMethod, "", "Provision", "Error",
            schemaPath, ex.Message);
        _run.AddErrorLog("Provisioning", "CreateList", $"{src.SiteUrl} :: {sourceList}", $"{dst.SiteUrl} :: {destinationList}", ex);
        Log($"Provisioning error: {ex.Message}", MigratorLogLevel.Error);
        return ("Error", null);
    }

    /// <summary>
    /// Creates a list with the typed title and the URL-safe URL in one step (CSOM lets the URL
    /// be set at creation, so no later rename is needed when naming is enforced).
    /// </summary>
    private SpListInfo CreateList(SpConnection c, string title, int templateType, Guid templateFeatureId, bool isLibrary,
        string description)
    {
        var info = new ListCreationInformation
        {
            Title = title.Trim(),
            TemplateType = templateType,
            Description = description ?? ""
        };
        if (templateFeatureId != Guid.Empty) info.TemplateFeatureId = templateFeatureId;
        if (_s.Naming.EnforceDestinationNaming)
        {
            var urlName = GetUrlSafeName(title);
            if (!string.IsNullOrWhiteSpace(urlName)) info.Url = isLibrary ? urlName : "Lists/" + urlName;
        }
        var list = c.Web.Lists.Add(info);
        c.Context.Load(list, l => l.Id);
        c.ExecuteQuery();
        Log($"Created list '{title}'{(info.Url != null ? $" at URL '{info.Url}'" : "")} (template {templateType}).", MigratorLogLevel.Success);
        return TryGetList(c, list.Id.ToString()) ?? throw new InvalidOperationException("The new list could not be read back.");
    }

    // ======================================================================= columns (shared with Modernize)

    private sealed class ColumnPlan
    {
        public SpFieldInfo Field { get; init; }
        public bool Include { get; init; }
        public string Reason { get; init; }
        /// <summary>False for routine exclusions (hidden/system) that would only add noise to the log.</summary>
        public bool LogSkip { get; init; }
    }

    private List<ColumnPlan> PlanColumns(IEnumerable<SpFieldInfo> sourceFields, HashSet<string> excludedTypes, string purpose)
    {
        var plan = new List<ColumnPlan>();
        foreach (var f in sourceFields)
        {
            string reason = "";
            var logSkip = false;
            if (f.Hidden) reason = "Hidden source field.";
            else if (f.ReadOnlyField) reason = "Read-only source field.";
            else if (_systemFields.Contains(f.InternalName)) reason = "System field.";
            else if (f.InternalName.StartsWith("_")) reason = "Internal/system-style field.";
            else if (f.InternalName == "Title") reason = "Title already exists on a new list.";
            else if (excludedTypes.Contains(f.TypeAsString))
            {
                reason = $"Unsupported type '{f.TypeAsString}' for {purpose}.";
                logSkip = true;
            }
            plan.Add(new ColumnPlan { Field = f, Include = reason.Length == 0, Reason = reason, LogSkip = logSkip });
        }
        return plan;
    }

    /// <summary>
    /// Strips identity/binding attributes that are only valid where the
    /// column came from, gives it a new ID, unhides it, and KEEPS Name/StaticName so metadata
    /// replay (which matches by InternalName) works on the new list.
    /// </summary>
    private static string GetPortableFieldXml(SpFieldInfo field, string fieldGroup)
    {
        var doc = new XmlDocument();
        doc.LoadXml(field.SchemaXml);
        var node = doc.DocumentElement!;
        foreach (var attr in new[]
                 {
                     "SourceID", "WebId", "ListId", "List", "Version", "ColName", "RowOrdinal", "FromBaseType",
                     "RelationshipDeleteBehavior", "Aggregation", "Customization", "ReadOnly", "ReadOnlyEnforced"
                 })
        {
            if (node.HasAttribute(attr)) node.RemoveAttribute(attr);
        }
        node.SetAttribute("ID", "{" + Guid.NewGuid() + "}");
        if (node.HasAttribute("Hidden")) node.SetAttribute("Hidden", "FALSE");
        if (!string.IsNullOrWhiteSpace(fieldGroup)) node.SetAttribute("Group", fieldGroup);
        return node.OuterXml;
    }

    /// <summary>Adds planned columns with AddFieldAsXml (InternalName hint honoured). Idempotent.</summary>
    private (int Created, int Existing, int Failed) AddColumns(SpConnection dst, SpListInfo dstList, List<ColumnPlan> toCreate,
        string sourceSiteUrl, string sourceList, string destinationList, string method, string fieldGroup)
    {
        var destNames = new HashSet<string>(GetFields(dst, dstList.List).Select(f => f.InternalName), StringComparer.OrdinalIgnoreCase);
        int created = 0, existing = 0, failed = 0, i = 0;

        foreach (var c in toCreate)
        {
            i++;
            if (CancelRequested)
            {
                Log("Cancellation requested. Stopping column creation.", MigratorLogLevel.Warn);
                break;
            }
            _run.Progress("Creating list columns", i, toCreate.Count, c.Field.InternalName);

            if (destNames.Contains(c.Field.InternalName))
            {
                Log($"  Column already exists, skipping: {c.Field.InternalName}");
                _run.AddProvisionLog(sourceSiteUrl, sourceList, dst.SiteUrl, destinationList, method, "",
                    "CreateColumn", "AlreadyExists", "", $"Column '{c.Field.InternalName}' already present on destination.");
                existing++;
                continue;
            }

            try
            {
                var xml = GetPortableFieldXml(c.Field, fieldGroup);
                dstList.List.Fields.AddFieldAsXml(xml, false, AddFieldOptions.AddFieldInternalNameHint);
                dst.ExecuteQuery();
                destNames.Add(c.Field.InternalName);
                Log($"  Column created: {c.Field.InternalName} ({c.Field.TypeAsString})", MigratorLogLevel.Success);
                _run.AddProvisionLog(sourceSiteUrl, sourceList, dst.SiteUrl, destinationList, method, "",
                    "CreateColumn", "Created", "", $"Column '{c.Field.InternalName}' ({c.Field.TypeAsString}) created from source schema.");
                created++;
            }
            catch (Exception ex)
            {
                Log($"  Failed to create column '{c.Field.InternalName}' ({c.Field.TypeAsString}): {ex.Message}", MigratorLogLevel.Error);
                _run.AddProvisionLog(sourceSiteUrl, sourceList, dst.SiteUrl, destinationList, method, "",
                    "CreateColumn", "Failed", "", $"Column '{c.Field.InternalName}' failed: {ex.Message}");
                _run.AddErrorLog(method, "AddFieldAsXml", $"{sourceSiteUrl} :: {sourceList} :: {c.Field.InternalName}",
                    $"{dst.SiteUrl} :: {destinationList}", ex);
                failed++;
            }
        }

        Log($"Column creation complete. Created: {created} | Already existed: {existing} | Failed: {failed}", MigratorLogLevel.Success);
        if (toCreate.Count > 0 && created == 0 && existing == 0)
        {
            Log("NO columns were created. The destination list exists but has no source metadata - review ErrorLog.csv before copying items.",
                MigratorLogLevel.Error);
        }
        return (created, existing, failed);
    }

    // ======================================================================= content types

    /// <summary>
    /// Binds each visible source list content type to the destination list when its parent
    /// content type exists in the destination site. Returns the number of warnings.
    /// </summary>
    private int CopyListContentTypes(SpConnection src, SpListInfo srcList, SpConnection dst, SpListInfo dstList,
        string destinationList)
    {
        var warnings = 0;
        try
        {
            var sourceTypes = srcList.List.ContentTypes;
            src.Context.Load(sourceTypes, cts => cts.Include(ct => ct.Name, ct => ct.Hidden, ct => ct.Parent.Id));
            var destTypes = dstList.List.ContentTypes;
            dst.Context.Load(destTypes, cts => cts.Include(ct => ct.Name));
            src.ExecuteQuery();
            dst.ExecuteQuery();

            var destNames = new HashSet<string>(destTypes.Select(ct => ct.Name), StringComparer.OrdinalIgnoreCase);
            foreach (var ct in sourceTypes)
            {
                if (ct.Hidden || ct.Name == "Folder" || destNames.Contains(ct.Name)) continue;
                var parentId = ct.Parent.Id.StringValue;
                try
                {
                    var siteType = dst.Web.AvailableContentTypes.GetById(parentId);
                    dst.Context.Load(siteType, t => t.Name);
                    dst.ExecuteQuery();
                    if (siteType.ServerObjectIsNull == true)
                    {
                        throw new InvalidOperationException($"Parent content type {parentId} does not exist in the destination site.");
                    }
                    dstList.List.ContentTypes.AddExistingContentType(siteType);
                    dst.ExecuteQuery();
                    destNames.Add(ct.Name);
                    Log($"  Content type added: {ct.Name}", MigratorLogLevel.Success);
                    _run.AddProvisionLog(src.SiteUrl, srcList.Title, dst.SiteUrl, destinationList, NativeMethod, "",
                        "AddContentType", "Created", "", $"Content type '{ct.Name}' bound to the destination list.");
                }
                catch (Exception ex)
                {
                    warnings++;
                    Log($"  Content type '{ct.Name}' not added: {ex.Message}", MigratorLogLevel.Warn);
                    _run.AddProvisionLog(src.SiteUrl, srcList.Title, dst.SiteUrl, destinationList, NativeMethod, "",
                        "AddContentType", "Skipped", "", $"Content type '{ct.Name}' not added: {ex.Message}");
                }
            }
        }
        catch (Exception ex)
        {
            warnings++;
            Log($"Could not read content types: {ex.Message}", MigratorLogLevel.Warn);
        }
        return warnings;
    }

    // ======================================================================= schema snapshot

    /// <summary>Writes the source list's settings, column SchemaXml and views to an XML file for review.</summary>
    private string WriteSchemaSnapshot(SpConnection src, SpListInfo srcList, List<SpFieldInfo> fields)
    {
        var views = srcList.List.Views;
        src.Context.Load(views, vs => vs.Include(v => v.Title, v => v.ListViewXml));
        src.ExecuteQuery();

        var doc = new XmlDocument();
        var root = doc.CreateElement("ListSchema");
        doc.AppendChild(root);
        root.SetAttribute("SourceSite", src.SiteUrl);
        root.SetAttribute("Title", srcList.Title);
        root.SetAttribute("Url", srcList.RootFolderUrl);
        root.SetAttribute("BaseTemplate", srcList.BaseTemplate.ToString());
        root.SetAttribute("BaseType", srcList.BaseType.ToString());
        root.SetAttribute("EnableVersioning", srcList.EnableVersioning.ToString());
        root.SetAttribute("MajorVersionLimit", srcList.MajorVersionLimit.ToString());
        root.SetAttribute("EnableAttachments", srcList.EnableAttachments.ToString());
        root.SetAttribute("ContentTypesEnabled", srcList.ContentTypesEnabled.ToString());
        root.SetAttribute("Extracted", DateTime.Now.ToString("s"));

        var fieldsNode = doc.CreateElement("Fields");
        root.AppendChild(fieldsNode);
        foreach (var f in fields)
        {
            try
            {
                var frag = doc.CreateDocumentFragment();
                frag.InnerXml = f.SchemaXml;
                fieldsNode.AppendChild(frag);
            }
            catch (XmlException)
            {
                var c = doc.CreateComment($" Field {f.InternalName}: SchemaXml could not be parsed ");
                fieldsNode.AppendChild(c);
            }
        }

        var viewsNode = doc.CreateElement("Views");
        root.AppendChild(viewsNode);
        foreach (var v in views)
        {
            try
            {
                var frag = doc.CreateDocumentFragment();
                frag.InnerXml = v.ListViewXml;
                viewsNode.AppendChild(frag);
            }
            catch (XmlException)
            {
                viewsNode.AppendChild(doc.CreateComment($" View {v.Title}: XML could not be parsed "));
            }
        }

        var path = _run.RunPath($"ProvisioningSchema_{RunContext.SafeFileName(srcList.Title)}.xml");
        System.IO.Directory.CreateDirectory(_run.RunFolder);
        doc.Save(path);
        return path;
    }

    // ======================================================================= naming (v2.3.1)

    /// <summary>
    /// Forces the list to the typed TITLE and the matching URL segment.
    /// URL first (RootFolder.MoveTo), then title. Idempotent. Returns the title to use downstream.
    /// </summary>
    private string SetListTitleAndUrl(SpConnection c, string currentListTitle, string desiredTitle, string modeName)
    {
        if (!_s.Naming.EnforceDestinationNaming) return currentListTitle;

        desiredTitle = desiredTitle.Trim();
        var desiredUrlName = GetUrlSafeName(desiredTitle);
        if (string.IsNullOrWhiteSpace(desiredUrlName))
        {
            Log($"Destination name '{desiredTitle}' produced an empty URL segment. Leaving the list as '{currentListTitle}'.", MigratorLogLevel.Warn);
            return currentListTitle;
        }

        var list = TryGetList(c, currentListTitle);
        if (list == null)
        {
            Log($"Could not load list '{currentListTitle}' to apply destination naming.", MigratorLogLevel.Error);
            return currentListTitle;
        }

        var currentServerRelative = list.RootFolderUrl ?? "";
        var currentUrlLeaf = GetLeafName(currentServerRelative);
        var isUnderLists = Regex.IsMatch(currentServerRelative, "/Lists/[^/]+/?$", RegexOptions.IgnoreCase);
        var targetSiteRelative = isUnderLists ? "Lists/" + desiredUrlName : desiredUrlName;

        // ---- 1. URL
        if (!string.Equals(currentUrlLeaf, desiredUrlName, StringComparison.OrdinalIgnoreCase))
        {
            try
            {
                list.List.RootFolder.MoveTo(targetSiteRelative);
                c.ExecuteQuery();
                Log($"Destination URL set: '{currentUrlLeaf}' -> '{desiredUrlName}'", MigratorLogLevel.Success);
                _run.AddProvisionLog("", "", c.SiteUrl, desiredTitle, "Naming", "", "SetListUrl", "Changed", "",
                    $"URL segment changed from '{currentUrlLeaf}' to '{desiredUrlName}' (target: {targetSiteRelative}).");
            }
            catch (Exception ex)
            {
                Log($"Could not change the list URL to '{desiredUrlName}': {ex.Message}", MigratorLogLevel.Error);
                Log($"  The display title will still be corrected, but the URL will remain '{currentUrlLeaf}'. " +
                    "A list or folder may already exist at the target URL.", MigratorLogLevel.Warn);
                _run.AddProvisionLog("", "", c.SiteUrl, desiredTitle, "Naming", "", "SetListUrl", "Failed", "",
                    $"URL rename to '{desiredUrlName}' failed: {ex.Message}");
                _run.AddErrorLog(modeName, "RootFolder.MoveTo", currentServerRelative, targetSiteRelative, ex);
                if (_s.Naming.UrlRenameFailureIsFatal) throw;
            }
        }
        else
        {
            Log($"Destination URL already correct: '{desiredUrlName}'", MigratorLogLevel.Debug);
        }

        // ---- 2. Title
        if (!string.Equals(list.Title, desiredTitle, StringComparison.Ordinal))
        {
            try
            {
                var l = c.Web.Lists.GetById(list.Id);
                l.Title = desiredTitle;
                l.Update();
                c.ExecuteQuery();
                Log($"Destination title set: '{list.Title}' -> '{desiredTitle}'", MigratorLogLevel.Success);
                _run.AddProvisionLog("", "", c.SiteUrl, desiredTitle, "Naming", "", "SetListTitle", "Changed", "",
                    $"Title changed from '{list.Title}' to '{desiredTitle}'.");
            }
            catch (Exception ex)
            {
                Log($"Could not set the list title to '{desiredTitle}': {ex.Message}", MigratorLogLevel.Error);
                _run.AddErrorLog(modeName, "SetListTitle", currentListTitle, desiredTitle, ex);
                return currentListTitle;
            }
        }
        else
        {
            Log($"Destination title already correct: '{desiredTitle}'", MigratorLogLevel.Debug);
        }
        return desiredTitle;
    }

    // ======================================================================= views (v2.2.6)

    /// <summary>
    /// Rebuilds public (optionally personal) views, keeping only columns that
    /// exist on the destination and dropping a query that references a missing column.
    /// </summary>
    private void CopyListViews(SpConnection src, SpListInfo srcList, SpConnection dst, SpListInfo dstList)
    {
        int createdCount = 0, skipped = 0, failed = 0;
        if (!_s.Provisioning.EnableViewReplication)
        {
            Log("View replication is disabled in Settings. Skipping.");
            return;
        }
        Log($"Replicating views from '{srcList.Title}' to '{dstList.Title}'...");

        ViewCollection sourceViews;
        try
        {
            sourceViews = srcList.List.Views;
            src.Context.Load(sourceViews, vs => vs.Include(
                v => v.Title, v => v.Hidden, v => v.PersonalView, v => v.ViewFields, v => v.ViewQuery,
                v => v.RowLimit, v => v.Paged, v => v.DefaultView));
            src.ExecuteQuery();
        }
        catch (Exception ex)
        {
            Log($"Could not read source views: {ex.Message}", MigratorLogLevel.Warn);
            return;
        }
        if (sourceViews.Count == 0)
        {
            Log("No views found on the source list.");
            return;
        }

        HashSet<string> destFieldNames;
        try
        {
            destFieldNames = new HashSet<string>(GetFields(dst, dstList.List).Select(f => f.InternalName), StringComparer.OrdinalIgnoreCase);
        }
        catch (Exception ex)
        {
            Log($"Could not read destination fields; view replication aborted: {ex.Message}", MigratorLogLevel.Warn);
            return;
        }

        var existingViews = new Dictionary<string, View>(StringComparer.OrdinalIgnoreCase);
        try
        {
            var dv = dstList.List.Views;
            dst.Context.Load(dv, vs => vs.Include(v => v.Title, v => v.Id));
            dst.ExecuteQuery();
            foreach (var v in dv) existingViews[v.Title] = v;
        }
        catch { }

        foreach (var view in sourceViews)
        {
            if (CancelRequested)
            {
                Log("Cancellation requested. Stopping view replication.", MigratorLogLevel.Warn);
                break;
            }
            var title = view.Title;
            if (string.IsNullOrWhiteSpace(title)) continue;

            if (view.Hidden)
            {
                Log($"  Skipping hidden view: {title}", MigratorLogLevel.Debug);
                skipped++;
                continue;
            }
            if (view.PersonalView && !_s.Provisioning.ReplicatePersonalViews)
            {
                Log($"  Skipping personal view: {title}", MigratorLogLevel.Debug);
                skipped++;
                continue;
            }

            if (existingViews.TryGetValue(title, out var existing))
            {
                if (!_s.Provisioning.OverwriteExistingViews)
                {
                    Log($"  View already exists, skipping: {title}");
                    _run.AddProvisionLog(src.SiteUrl, srcList.Title, dst.SiteUrl, dstList.Title, "NativeView", "",
                        "CreateView", "AlreadyExists", "", $"View '{title}' already present on destination.");
                    skipped++;
                    continue;
                }
                try
                {
                    existing.DeleteObject();
                    dst.ExecuteQuery();
                    Log($"  Removed existing view for recreate: {title}", MigratorLogLevel.Warn);
                }
                catch (Exception ex)
                {
                    Log($"  Could not remove existing view '{title}': {ex.Message}", MigratorLogLevel.Warn);
                    skipped++;
                    continue;
                }
            }

            var kept = new List<string>();
            var omitted = new List<string>();
            foreach (var vf in view.ViewFields)
            {
                if (string.IsNullOrWhiteSpace(vf)) continue;
                if (destFieldNames.Contains(vf)) kept.Add(vf); else omitted.Add(vf);
            }
            if (omitted.Count > 0)
            {
                Log($"  View '{title}': omitting {omitted.Count} missing column(s): {string.Join(", ", omitted)}", MigratorLogLevel.Warn);
            }
            if (kept.Count == 0)
            {
                Log($"  View '{title}' has no usable columns on the destination. Skipping.", MigratorLogLevel.Warn);
                _run.AddProvisionLog(src.SiteUrl, srcList.Title, dst.SiteUrl, dstList.Title, "NativeView", "",
                    "CreateView", "Skipped", "", $"View '{title}' skipped: none of its columns exist on the destination.");
                skipped++;
                continue;
            }

            var query = view.ViewQuery ?? "";
            var queryDropped = false;
            foreach (Match m in Regex.Matches(query, "<FieldRef[^>]*Name\\s*=\\s*['\"]([^'\"]+)['\"]"))
            {
                if (!destFieldNames.Contains(m.Groups[1].Value))
                {
                    queryDropped = true;
                    break;
                }
            }
            if (queryDropped)
            {
                Log($"  View '{title}': filter/sort references a missing column - query dropped.", MigratorLogLevel.Warn);
                query = "";
            }

            try
            {
                var info = new ViewCreationInformation
                {
                    Title = title,
                    ViewFields = kept.ToArray(),
                    Paged = view.Paged,
                    SetAsDefaultView = view.DefaultView,
                    PersonalView = view.PersonalView && _s.Provisioning.ReplicatePersonalViews,
                    ViewTypeKind = ViewType.Html
                };
                if (!string.IsNullOrWhiteSpace(query)) info.Query = query;
                if (view.RowLimit > 0) info.RowLimit = view.RowLimit;
                dstList.List.Views.Add(info);
                dst.ExecuteQuery();

                var status = omitted.Count > 0 || queryDropped ? "CreatedWithOmittedFields" : "Created";
                var note = $"View '{title}' created.";
                if (omitted.Count > 0) note += $" Omitted columns: {string.Join(", ", omitted)}.";
                if (queryDropped) note += " Filter/sort query dropped (referenced a missing column).";
                _run.AddProvisionLog(src.SiteUrl, srcList.Title, dst.SiteUrl, dstList.Title, "NativeView", "",
                    "CreateView", status, "", note);
                Log($"  View created: {title}{(omitted.Count > 0 ? " (with omitted columns)" : "")}", MigratorLogLevel.Success);
                createdCount++;
            }
            catch (Exception ex)
            {
                Log($"  Failed to create view '{title}': {ex.Message}", MigratorLogLevel.Error);
                _run.AddProvisionLog(src.SiteUrl, srcList.Title, dst.SiteUrl, dstList.Title, "NativeView", "",
                    "CreateView", "Failed", "", $"View '{title}' failed: {ex.Message}");
                _run.AddErrorLog("ViewReplication", "AddView", $"{src.SiteUrl} :: {srcList.Title} :: {title}",
                    $"{dst.SiteUrl} :: {dstList.Title}", ex);
                failed++;
            }
        }
        Log($"View replication complete. Created: {createdCount} | Skipped: {skipped} | Failed: {failed}", MigratorLogLevel.Success);
    }
}
