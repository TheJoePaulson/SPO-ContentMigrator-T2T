using Microsoft.SharePoint.Client;
using ContentMigrator.Infrastructure;
using ContentMigrator.SharePoint;

namespace ContentMigrator.Engine;

public sealed partial class MigrationEngine
{
    private const string ModernizeMethod = "Modernize";

    /// <summary>GUI "ModernizeList" (Operation 10): new modern list, columns, views, then items.</summary>
    public void RunModernize(string sourceSiteUrl, string sourceList, string destinationSiteUrl, string destinationList)
    {
        if (!_s.Modernize.EnableModernizeList)
        {
            Log("Modernize List is disabled in Settings.", MigratorLogLevel.Warn);
            return;
        }
        var src = Connect(sourceSiteUrl);
        var dst = Connect(destinationSiteUrl, TenantSide.Destination);
        LogPreflight("ModernizeList", "Copy", src.SiteUrl, sourceList, dst.SiteUrl, destinationList, 0,
            "Creates a NEW modern (GenericList) list, recreates supported columns and views, then copies items.");

        var (status, dstList, srcList) = CreateModernList(src, sourceList, dst, destinationList);
        if (status is not ("Created" or "AlreadyExists"))
        {
            Log($"Modernize did not produce a usable destination list (status: {status}). Item copy skipped.", MigratorLogLevel.Warn);
            Log($"Modernize List completed with status: {status}");
            return;
        }

        WriteFieldCompatibilityReport(src, srcList, dst, dstList, "ModernizeList");
        if (!_s.Modernize.CopyItems)
        {
            Log("Copy items after modernizing is off. Schema and views created; no items copied.");
        }
        else
        {
            CopyListItems(src, srcList, dst, dstList);
        }
        Log($"Modernize List completed with status: {status}");
    }

    /// <summary>
    /// Creates a GenericList (template 100) with the modern
    /// experience, re-adds the source's supported columns by InternalName, and rebuilds views.
    /// Reuses an existing destination and adds only missing columns. Honours review-only.
    /// </summary>
    private (string Status, SpListInfo Destination, SpListInfo Source) CreateModernList(SpConnection src, string sourceList,
        SpConnection dst, string destinationList)
    {
        if (!CheckSourceAndDestinationDiffer(src.SiteUrl, sourceList, dst.SiteUrl, destinationList, ModernizeMethod))
        {
            return ("Error", null, null);
        }

        var srcList = TryGetList(src, sourceList);
        if (srcList == null)
        {
            Log($"Source list '{sourceList}' was not found. Cannot modernize.", MigratorLogLevel.Error);
            _run.AddProvisionLog(src.SiteUrl, sourceList, dst.SiteUrl, destinationList, ModernizeMethod, "", "Modernize",
                "Error", "", "Source list not found.");
            return ("Error", null, null);
        }

        List<SpFieldInfo> sourceFields;
        try
        {
            sourceFields = GetFields(src, srcList.List);
        }
        catch (Exception ex)
        {
            Log($"Could not read source fields: {ex.Message}", MigratorLogLevel.Error);
            return ("Error", null, srcList);
        }

        var excluded = new HashSet<string>(_s.Modernize.ExcludedFieldTypes, StringComparer.OrdinalIgnoreCase);
        var plan = PlanColumns(sourceFields, excluded, "modernization");
        var toCreate = plan.Where(p => p.Include).ToList();
        var toSkip = plan.Where(p => !p.Include).ToList();

        Log($"Modernize plan for '{sourceList}' -> '{destinationList}': {toCreate.Count} column(s) to create, {toSkip.Count} skipped.");
        foreach (var skip in toSkip)
        {
            if (_s.Reports.LogMetadataSkipped)
            {
                Log($"  Skipping column '{skip.Field.InternalName}' ({skip.Field.TypeAsString}): {skip.Reason}", MigratorLogLevel.Debug);
            }
            _run.AddProvisionLog(src.SiteUrl, sourceList, dst.SiteUrl, destinationList, ModernizeMethod, "", "SkipColumn",
                "Skipped", "", $"Column '{skip.Field.InternalName}' ({skip.Field.TypeAsString}) skipped: {skip.Reason}");
        }
        foreach (var c in toCreate)
        {
            Log($"  Will create column: {c.Field.InternalName} ({c.Field.TypeAsString}) - '{c.Field.Title}'");
        }

        if (_opt.ReviewOnly)
        {
            Log("Review-only mode is enabled. The modern list was NOT created; the column plan above was logged only.", MigratorLogLevel.Warn);
            foreach (var c in toCreate)
            {
                _run.AddProvisionLog(src.SiteUrl, sourceList, dst.SiteUrl, destinationList, ModernizeMethod, "", "PlanColumn",
                    "ReviewOnly", "", $"Would create column '{c.Field.InternalName}' ({c.Field.TypeAsString}).");
            }
            return ("ReviewOnly", null, srcList);
        }

        // ---- create (or reuse) the destination list
        var dstList = TryGetList(dst, destinationList);
        var existed = dstList != null;
        if (existed)
        {
            Log($"Destination list '{destinationList}' already exists. Reusing it and adding only missing columns.");
        }
        else
        {
            try
            {
                dstList = CreateList(dst, destinationList, (int)ListTemplateType.GenericList, Guid.Empty, false, srcList.Description);
                SetListTitleAndUrl(dst, dstList.Title, destinationList, ModernizeMethod);
            }
            catch (Exception ex)
            {
                Log($"Failed to create modern list '{destinationList}': {ex.Message}", MigratorLogLevel.Error);
                _run.AddProvisionLog(src.SiteUrl, sourceList, dst.SiteUrl, destinationList, ModernizeMethod, "", "Modernize",
                    "Error", "", $"List creation failed: {ex.Message}");
                _run.AddErrorLog(ModernizeMethod, "CreateList", $"{src.SiteUrl} :: {sourceList}", $"{dst.SiteUrl} :: {destinationList}", ex);
                return ("Error", null, srcList);
            }
        }

        // ---- modern experience / versioning / attachments
        try
        {
            var l = dstList.List;
            l.ListExperienceOptions = ListExperience.NewExperience;
            l.EnableAttachments = _s.Modernize.EnableAttachments;
            if (!existed) l.EnableVersioning = _s.Modernize.EnableVersioning;
            l.Update();
            dst.ExecuteQuery();
            Log($"List experience set to 'NewExperience'; attachments enabled: {_s.Modernize.EnableAttachments}.", MigratorLogLevel.Success);
        }
        catch (Exception ex)
        {
            Log($"Could not apply list experience/attachment settings: {ex.Message}", MigratorLogLevel.Warn);
        }

        // Re-read so EnableAttachments etc. reflect what was just applied.
        dstList = TryGetList(dst, dstList.Id.ToString()) ?? dstList;

        var group = string.IsNullOrWhiteSpace(_s.Modernize.FieldGroup) ? null : _s.Modernize.FieldGroup;
        var (created, already, failed) = AddColumns(dst, dstList, toCreate, src.SiteUrl, sourceList, destinationList,
            ModernizeMethod, group);

        try
        {
            CopyListViews(src, srcList, dst, dstList);
        }
        catch (Exception ex)
        {
            Log($"View replication failed: {ex.Message}", MigratorLogLevel.Warn);
        }

        var status = existed ? "AlreadyExists" : "Created";
        _run.AddProvisionLog(src.SiteUrl, sourceList, dst.SiteUrl, destinationList, ModernizeMethod, "", "Modernize", status, "",
            $"Modern list ready. ColumnsCreated={created}; AlreadyExisted={already}; Failed={failed}; Skipped={toSkip.Count}.");
        return (status, dstList, srcList);
    }
}
