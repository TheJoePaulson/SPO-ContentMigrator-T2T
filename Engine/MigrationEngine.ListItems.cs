using System.Diagnostics;
using Microsoft.SharePoint.Client;
using ContentMigrator.Infrastructure;
using ContentMigrator.SharePoint;
using IOFile = System.IO.File;
using IODirectory = System.IO.Directory;
using IOPath = System.IO.Path;

namespace ContentMigrator.Engine;

public sealed partial class MigrationEngine
{
    private sealed class AttachmentTotals
    {
        public int Copied;
        public int Skipped;
        public int Failed;
    }

    // ======================================================================= mode 6

    /// <summary>GUI "ListItemCopy": connect, ensure the destination, then copy every item.</summary>
    public void RunListItemCopy(string sourceSiteUrl, string sourceList, string destinationSiteUrl, string destinationList)
    {
        var src = Connect(sourceSiteUrl);
        var dst = Connect(destinationSiteUrl, TenantSide.Destination);
        var dstList = ConfirmDestinationOrProvision(src, sourceList, dst, destinationList);
        if (dstList == null)
        {
            Log("Destination not available. Operation aborted.", MigratorLogLevel.Warn);
            return;
        }
        CopyListItems(src, GetListOrThrow(src, sourceList), dst, dstList);
    }

    /// <summary>Copies every item of a list into another list.</summary>
    private void CopyListItems(SpConnection src, SpListInfo srcList, SpConnection dst, SpListInfo dstList)
    {
        // v2.3.1: copying a list into itself would duplicate every item. Stop first.
        if (!CheckSourceAndDestinationDiffer(src.SiteUrl, srcList.Title, dst.SiteUrl, dstList.Title, "ListItemCopy")) return;

        var fieldMap = GetWritableFieldMap(src, srcList, dst, dstList);
        WriteFieldCompatibilityReport(src, srcList, dst, dstList, "ListItemCopy");
        Log($"Writable matching field count: {fieldMap.Count}");

        // v2.3.0: decide ONCE whether attachments can be copied for this operation.
        var copyAttachments = AttachmentsSupported(dst, dstList);
        var totals = new AttachmentTotals();

        var items = GetAllItems(src, srcList.List);
        LogPreflight("ListItemCopy", "Copy", src.SiteUrl, srcList.Title, dst.SiteUrl, dstList.Title, items.Count,
            "List item copy assumes matching source/destination list schemas.");

        var count = 0;
        foreach (var item in items)
        {
            count++;
            if (CancelRequested)
            {
                Log($"Cancellation requested. Stopping list copy after {count - 1} item(s).", MigratorLogLevel.Warn);
                break;
            }
            _run.Progress("Copying SharePoint list items", count, items.Count, $"Item ID {item.Id}");

            var timer = Stopwatch.StartNew();
            var sourceRef = $"{src.SiteUrl} | {srcList.Title} | ItemId={item.Id}";
            var destinationRef = $"{dst.SiteUrl} | {dstList.Title}";
            try
            {
                var (values, _, skipped) = BuildMetadataValues(item, fieldMap, sourceRef, destinationRef,
                    logCopied: false, logSkipped: _s.Reports.LogMetadataSkipped);
                var resolved = ResolveValuesForWrite(dst, values, sourceRef);

                var newItem = dstList.List.AddItem(new ListItemCreationInformation());
                foreach (var kv in resolved) newItem[kv.Key] = kv.Value;
                newItem.Update();
                dst.Context.Load(newItem, i => i.Id);
                dst.ExecuteQuery();

                var newRef = $"{dst.SiteUrl} | {dstList.Title} | ItemId={newItem.Id}";
                var systemStatus = ApplySystemMetadata(dst, newItem, item, sourceRef, newRef);

                // v2.3.0: attachments ride along after the item exists (they need its ID).
                var attachStatus = "None";
                if (copyAttachments)
                {
                    var r = CopyItemAttachments(src, srcList, item, dst, dstList, newItem.Id);
                    totals.Copied += r.Copied;
                    totals.Skipped += r.Skipped;
                    totals.Failed += r.Failed;
                    if (r.Copied + r.Skipped + r.Failed > 0)
                        attachStatus = $"Copied:{r.Copied};Skipped:{r.Skipped};Failed:{r.Failed}";
                    if (r.Failed > 0 && _s.Attachments.AttachmentFailureIsItemFailure)
                        throw new InvalidOperationException($"Item data copied, but {r.Failed} attachment(s) failed. See AttachmentLog.csv.");
                }

                _run.AddOperationLog("ListItemCopy", "Copy", src.SiteUrl, sourceRef, dst.SiteUrl, newRef, "ListItem",
                    "Success", "", $"FieldsCopied:{resolved.Count};Skipped:{skipped};System:{systemStatus};Attachments:{attachStatus}",
                    timer.Elapsed.TotalSeconds);
            }
            catch (Exception ex)
            {
                _run.AddOperationLog("ListItemCopy", "Copy", src.SiteUrl, sourceRef, dst.SiteUrl, destinationRef, "ListItem",
                    "Error", ex.Message, "NotApplied", timer.Elapsed.TotalSeconds);
                _run.AddErrorLog("ListItemCopy", "Copy", sourceRef, destinationRef, ex);
                Log($"ERROR copying list item ID {item.Id}: {ex.Message}", MigratorLogLevel.Error);
            }

            if (count % 100 == 0)
            {
                Log($"Copied/evaluated {count} list items...");
                _run.ExportLogs();
            }
        }

        // v2.3.0: tidy the staging root once the whole run is done.
        if (!_s.Attachments.KeepStagingFiles)
        {
            var stagingRoot = _run.RunPath("AttachmentStaging");
            try { if (IODirectory.Exists(stagingRoot)) IODirectory.Delete(stagingRoot, true); } catch { }
        }

        Log($"List item copy complete. Items evaluated: {count}", MigratorLogLevel.Success);
        if (copyAttachments)
        {
            Log($"Attachments - Copied: {totals.Copied} | Skipped: {totals.Skipped} | Failed: {totals.Failed}", MigratorLogLevel.Success);
        }
    }

    // ======================================================================= attachments (v2.3.0)

    /// <summary>Decides once per operation whether attachments can be copied.</summary>
    private bool AttachmentsSupported(SpConnection dst, SpListInfo dstList)
    {
        if (!_s.Attachments.EnableAttachmentCopy)
        {
            Log("Attachment copy is disabled in Settings. Item attachments will NOT be migrated.");
            return false;
        }
        if (dstList.IsLibrary)
        {
            Log("Destination is a document library; list item attachments do not apply.", MigratorLogLevel.Debug);
            return false;
        }
        if (!dstList.EnableAttachments)
        {
            Log($"Destination list '{dstList.Title}' has attachments DISABLED. Item attachments will NOT be migrated. " +
                "Enable attachments on the list, or turn on 'Enable attachments on the new list' in Settings > Modernize.",
                MigratorLogLevel.Warn);
            return false;
        }
        Log("Attachment copy is ENABLED for this operation.");
        return true;
    }

    /// <summary>
    /// Each attachment is downloaded to a per-item staging folder
    /// under the run folder and uploaded to the new destination item under the same name.
    /// Oversized files are skipped before download. Every attachment gets an AttachmentLog row.
    /// </summary>
    private AttachmentTotals CopyItemAttachments(SpConnection src, SpListInfo srcList, ListItem sourceItem,
        SpConnection dst, SpListInfo dstList, int destinationItemId)
    {
        var result = new AttachmentTotals();

        // Cheap gate: the "Attachments" field is already in FieldValues.
        if (sourceItem.FieldValues.TryGetValue("Attachments", out var hasAny) && hasAny is bool b && !b) return result;

        AttachmentCollection attachments;
        try
        {
            attachments = sourceItem.AttachmentFiles;
            src.Context.Load(attachments, a => a.Include(x => x.FileName, x => x.ServerRelativeUrl));
            src.ExecuteQuery();
        }
        catch (Exception ex)
        {
            Log($"  Attachment listing failed for item {sourceItem.Id}: {ex.Message}", MigratorLogLevel.Error);
            _run.AddAttachmentLog(src.SiteUrl, srcList.Title, sourceItem.Id, dst.SiteUrl, dstList.Title, destinationItemId,
                "", 0, "Failed", ex.Message);
            _run.AddErrorLog("AttachmentCopy", "ListAttachments", $"{src.SiteUrl} :: {srcList.Title} :: ItemId={sourceItem.Id}",
                $"{dst.SiteUrl} :: {dstList.Title} :: ItemId={destinationItemId}", ex);
            result.Failed++;
            return result;
        }
        if (attachments.Count == 0) return result;

        var stagingPath = IOPath.Combine(_run.RunPath("AttachmentStaging"), $"Item_{sourceItem.Id}");
        IODirectory.CreateDirectory(stagingPath);
        var maxBytes = _s.Attachments.MaxAttachmentSizeMB > 0 ? _s.Attachments.MaxAttachmentSizeMB * 1024L * 1024L : long.MaxValue;

        try
        {
            foreach (var attachment in attachments)
            {
                if (CancelRequested)
                {
                    Log("  Cancellation requested. Stopping attachment copy.", MigratorLogLevel.Warn);
                    break;
                }

                var name = attachment.FileName;
                long size = 0;
                try
                {
                    var sourceFile = src.Web.GetFileByServerRelativePath(ResourcePath.FromDecodedUrl(attachment.ServerRelativeUrl));
                    src.Context.Load(sourceFile, f => f.Length);
                    src.ExecuteQuery();
                    size = sourceFile.Length;

                    if (size > maxBytes)
                    {
                        var mb = Math.Round(size / 1024d / 1024d, 2);
                        Log($"  Skipping oversized attachment '{name}' ({mb} MB > {_s.Attachments.MaxAttachmentSizeMB} MB limit).", MigratorLogLevel.Warn);
                        _run.AddAttachmentLog(src.SiteUrl, srcList.Title, sourceItem.Id, dst.SiteUrl, dstList.Title,
                            destinationItemId, name, size, "Skipped",
                            $"Attachment is {mb} MB, above the {_s.Attachments.MaxAttachmentSizeMB} MB limit.");
                        result.Skipped++;
                        continue;
                    }

                    // Download to staging.
                    var localPath = IOPath.Combine(stagingPath, name);
                    var stream = sourceFile.OpenBinaryStream();
                    src.ExecuteQuery();
                    using (var source = stream.Value)
                    using (var local = IOFile.Create(localPath))
                    {
                        source.CopyTo(local);
                    }

                    // Upload to the new item.
                    using (var upload = IOFile.OpenRead(localPath))
                    {
                        var destinationItem = dstList.List.GetItemById(destinationItemId);
                        destinationItem.AttachmentFiles.Add(new AttachmentCreationInformation
                        {
                            FileName = name,
                            ContentStream = upload
                        });
                        dst.ExecuteQuery();
                    }

                    _run.AddAttachmentLog(src.SiteUrl, srcList.Title, sourceItem.Id, dst.SiteUrl, dstList.Title,
                        destinationItemId, name, size, "Copied", "");
                    result.Copied++;
                }
                catch (Exception ex)
                {
                    Log($"  Failed to copy attachment '{name}' to item {destinationItemId}: {ex.Message}", MigratorLogLevel.Error);
                    _run.AddAttachmentLog(src.SiteUrl, srcList.Title, sourceItem.Id, dst.SiteUrl, dstList.Title,
                        destinationItemId, name, size, "Failed", ex.Message);
                    _run.AddErrorLog("AttachmentCopy", "CopyAttachment",
                        $"{src.SiteUrl} :: {srcList.Title} :: ItemId={sourceItem.Id} :: {name}",
                        $"{dst.SiteUrl} :: {dstList.Title} :: ItemId={destinationItemId}", ex);
                    result.Failed++;
                }
            }

            if (result.Copied > 0)
            {
                Log($"  Item {sourceItem.Id} -> {destinationItemId} : {result.Copied} attachment(s) copied.", MigratorLogLevel.Success);
            }
        }
        finally
        {
            if (!_s.Attachments.KeepStagingFiles)
            {
                try { if (IODirectory.Exists(stagingPath)) IODirectory.Delete(stagingPath, true); } catch { }
            }
        }
        return result;
    }
}
