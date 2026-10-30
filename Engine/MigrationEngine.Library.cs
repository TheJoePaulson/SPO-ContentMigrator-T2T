using System.Diagnostics;
using System.Text.Json;
using Microsoft.SharePoint.Client;
using ContentMigrator.Infrastructure;
using ContentMigrator.SharePoint;

namespace ContentMigrator.Engine;

public sealed partial class MigrationEngine
{
    private readonly HashSet<string> _ensuredFolders = new(StringComparer.OrdinalIgnoreCase);

    private static readonly string[] RequiredCsvColumns =
    {
        "SourceSiteUrl", "SourceLibrary", "SourceServerRelativeUrl", "SourceRelativePathUnderRoot", "ItemType"
    };

    // ======================================================================= inventory

    /// <summary>Library inventory -> Reports\Inventory_&lt;library&gt;.csv</summary>
    public List<InventoryRecord> InventoryLibrary(SpConnection c, string libraryTitle, string outputFileName = null)
    {
        Log($"Starting inventory for library '{libraryTitle}'...");
        var lib = GetListOrThrow(c, libraryTitle);
        var libraryRoot = lib.RootFolderUrl;
        var fields = new[]
        {
            "ID", "UniqueId", "FileRef", "FileDirRef", "FileLeafRef", "FSObjType", "File_x0020_Size", "Created",
            "Modified", "Author", "Editor", "ContentType", "CheckoutUser"
        };
        var items = GetAllItems(c, lib.List, fields);
        var total = items.Count;
        var inventory = new List<InventoryRecord>();
        var counter = 0;

        foreach (var item in items)
        {
            counter++;
            var fv = item.FieldValues;
            var leaf = Get(fv, "FileLeafRef");
            _run.Progress($"Inventorying library '{libraryTitle}'", counter, total, leaf);

            var serverRelativeUrl = NormalizeServerRelativeUrl(Get(fv, "FileRef"));
            var itemType = Get(fv, "FSObjType") == "1" ? "Folder" : "File";
            if (itemType == "Folder" && !_s.Reports.IncludeFoldersInInventory) continue;
            if (itemType == "File" && !_s.Reports.IncludeFilesInInventory) continue;
            if (System.Text.RegularExpressions.Regex.IsMatch(serverRelativeUrl ?? "", "/Forms($|/)")) continue;

            var relativePath = GetRelativePathUnderRoot(serverRelativeUrl, libraryRoot);
            var folderRelativePath = itemType == "Folder" ? relativePath : GetParentFolderRelativePath(relativePath);

            inventory.Add(new InventoryRecord
            {
                SourceSiteUrl = c.SiteUrl,
                SourceLibrary = libraryTitle,
                SourceLibraryRoot = libraryRoot,
                ItemId = item.Id,
                UniqueId = Get(fv, "UniqueId"),
                ItemType = itemType,
                FileName = leaf,
                FileExtension = System.IO.Path.GetExtension(leaf ?? ""),
                SourceServerRelativeUrl = serverRelativeUrl,
                SourceRelativePathUnderRoot = relativePath,
                SourceFolderRelativePath = folderRelativePath,
                SourceFolderServerRelativeUrl = NormalizeServerRelativeUrl(Get(fv, "FileDirRef")),
                FileSizeBytes = Get(fv, "File_x0020_Size"),
                Created = fv.TryGetValue("Created", out var cr) ? cr as DateTime? : null,
                Modified = fv.TryGetValue("Modified", out var md) ? md as DateTime? : null,
                AuthorEmail = (fv.TryGetValue("Author", out var a) ? a as FieldUserValue : null)?.Email ?? "",
                EditorEmail = (fv.TryGetValue("Editor", out var e) ? e as FieldUserValue : null)?.Email ?? "",
                ContentType = Get(fv, "ContentType"),
                CheckoutUser = Get(fv, "CheckoutUser")
            });
            if (counter % 500 == 0) Log($"Inventoried {counter} items...");
        }

        outputFileName ??= $"Inventory_{RunContext.SafeFileName(libraryTitle)}.csv";
        var outputPath = _run.ReportPath(outputFileName);
        Csv.Write(outputPath, inventory);
        Log($"Inventory complete. Item count: {inventory.Count}", MigratorLogLevel.Success);
        Log($"Inventory file: {outputPath}", MigratorLogLevel.Success);
        return inventory;
    }

    private static string Get(Dictionary<string, object> fv, string name)
    {
        if (!fv.TryGetValue(name, out var v) || v == null) return "";
        return v switch
        {
            FieldUserValue u => u.LookupValue ?? "",
            FieldLookupValue l => l.LookupValue ?? "",
            _ => Convert.ToString(v, System.Globalization.CultureInfo.InvariantCulture)
        };
    }

    // ======================================================================= folders

    /// <summary>
    /// Creates each missing segment under the library root and returns
    /// the deepest folder's server-relative URL. Paths already ensured this run are cached, so
    /// each folder is checked (and logged) once instead of once per file.
    /// </summary>
    private string EnsureFolderPath(SpConnection c, string libraryRootServerRelativeUrl, string relativeFolderPath)
    {
        var root = NormalizeServerRelativeUrl(libraryRootServerRelativeUrl);
        if (string.IsNullOrWhiteSpace(relativeFolderPath)) return root;

        var current = root;
        foreach (var segment in relativeFolderPath.Trim('/').Split('/', StringSplitOptions.RemoveEmptyEntries))
        {
            var next = JoinUrlPath(current, segment);
            var key = c.SiteUrl + "|" + next;
            if (_ensuredFolders.Contains(key))
            {
                current = next;
                continue;
            }

            bool exists;
            try
            {
                var folder = c.Web.GetFolderByServerRelativePath(ResourcePath.FromDecodedUrl(next));
                c.Context.Load(folder, f => f.Exists);
                c.ExecuteQuery();
                exists = folder.Exists;
            }
            catch (ServerException)
            {
                exists = false;
            }

            if (exists)
            {
                _run.AddFolderLog(c.SiteUrl, next, "AlreadyExists", "");
            }
            else
            {
                try
                {
                    var parent = c.Web.GetFolderByServerRelativePath(ResourcePath.FromDecodedUrl(current));
                    parent.Folders.AddUsingPath(ResourcePath.FromDecodedUrl(next), new FolderCollectionAddParameters { Overwrite = false });
                    c.ExecuteQuery();
                    _run.AddFolderLog(c.SiteUrl, next, "Created", "");
                    Log($"Created folder: {next}", MigratorLogLevel.Success);
                }
                catch (Exception ex) when (ex.Message.Contains("exist", StringComparison.OrdinalIgnoreCase))
                {
                    _run.AddFolderLog(c.SiteUrl, next, "AlreadyExists", ex.Message);
                }
                catch (Exception ex)
                {
                    _run.AddFolderLog(c.SiteUrl, next, "Error", ex.Message);
                    _run.AddErrorLog("Folder", "EnsureFolder", "", next, ex);
                    throw;
                }
            }
            _ensuredFolders.Add(key);
            current = next;
        }
        return current;
    }

    // ======================================================================= file helpers

    /// <summary>True when the file exists AND has a list item (metadata is writable).</summary>
    private bool FileExists(SpConnection c, string fileServerRelativeUrl)
    {
        if (string.IsNullOrWhiteSpace(fileServerRelativeUrl)) return false;
        try
        {
            var file = c.Web.GetFileByServerRelativePath(ResourcePath.FromDecodedUrl(fileServerRelativeUrl));
            c.Context.Load(file, f => f.Exists);
            c.Context.Load(file.ListItemAllFields, i => i.Id);
            c.ExecuteQuery();
            return file.Exists && file.ListItemAllFields.Id > 0;
        }
        catch (ServerException)
        {
            return false;
        }
    }

    /// <summary>Polls until a copied/moved file is visible on the destination.</summary>
    private bool WaitForFile(SpConnection c, string fileServerRelativeUrl)
    {
        var max = _s.Copy.DestinationFileWaitMaxAttempts;
        for (var attempt = 1; attempt <= max; attempt++)
        {
            if (FileExists(c, fileServerRelativeUrl)) return true;
            Log($"Waiting for destination file to become available for metadata replay. Attempt {attempt}/{max} :: {fileServerRelativeUrl}",
                MigratorLogLevel.Debug);
            Thread.Sleep(TimeSpan.FromSeconds(_s.Copy.DestinationFileWaitSeconds));
        }
        return false;
    }

    /// <summary>Encodes a decoded absolute URL for APIs that take a URI string (copy jobs).</summary>
    private static string ToUriString(string decodedAbsoluteUrl) =>
        new Uri(decodedAbsoluteUrl.Replace("%", "%25").Replace("#", "%23")).AbsoluteUri;

    /// <summary>
    /// Copy-PnPFile / Move-PnPFile.
    ///   Same site:  MoveCopyUtil.CopyFileByPath / MoveFileByPath (keeps author/editor).
    ///   Cross site: Site.CreateCopyJobs (the SharePoint copy/move job API), honouring
    ///               AllowSchemaMismatch and AllowSmallerVersionLimitOnDestination.
    /// </summary>
    private void TransferFile(string operationType, SpConnection src, string sourceUrl, SpConnection dst,
        string destinationFolderUrl, string destinationFileUrl, bool isCrossSite)
    {
        var isMove = operationType == "Move";
        if (!isCrossSite)
        {
            var srcPath = ResourcePath.FromDecodedUrl(src.ToAbsoluteUrl(sourceUrl));
            var dstPath = ResourcePath.FromDecodedUrl(dst.ToAbsoluteUrl(destinationFileUrl));
            if (isMove)
            {
                MoveCopyUtil.MoveFileByPath(src.Context, srcPath, dstPath, _opt.OverwriteExisting,
                    new MoveCopyOptions { KeepBoth = false, RetainEditorAndModifiedOnMove = true });
            }
            else
            {
                MoveCopyUtil.CopyFileByPath(src.Context, srcPath, dstPath, _opt.OverwriteExisting,
                    new MoveCopyOptions { KeepBoth = false, ResetAuthorAndCreatedOnCopy = false });
            }
            src.ExecuteQuery();
            return;
        }

        var options = new CopyMigrationOptions
        {
            IsMoveMode = isMove,
            AllowSchemaMismatch = _s.Copy.AllowSchemaMismatch,
            AllowSmallerVersionLimitOnDestination = _s.Copy.AllowSmallerVersionLimitOnDestination,
            IgnoreVersionHistory = false,
            NameConflictBehavior = _opt.OverwriteExisting ? MigrationNameConflictBehavior.Replace : MigrationNameConflictBehavior.Fail
        };
        // CreateCopyJobs returns the job list directly (not a ClientResult); CSOM fills it
        // in when the query executes.
        IList<CopyMigrationInfo> jobs = src.Context.Site.CreateCopyJobs(
            new[] { ToUriString(src.ToAbsoluteUrl(sourceUrl)) },
            ToUriString(dst.ToAbsoluteUrl(destinationFolderUrl)),
            options);
        src.ExecuteQuery();
        if (jobs == null || jobs.Count == 0)
        {
            throw new InvalidOperationException("SharePoint did not return a copy job for this file.");
        }

        var job = jobs[0];
        var deadline = DateTime.UtcNow.AddMinutes(_s.Copy.CopyJobTimeoutMinutes);
        var errors = new List<string>();
        while (true)
        {
            var progress = src.Context.Site.GetCopyJobProgress(job);
            src.ExecuteQuery();
            foreach (var entry in progress.Value.Logs ?? Array.Empty<string>())
            {
                try
                {
                    using var doc = JsonDocument.Parse(entry);
                    var evt = doc.RootElement.TryGetProperty("Event", out var ev) ? ev.GetString() : "";
                    if (evt is "JobError" or "JobFatalError")
                    {
                        errors.Add(doc.RootElement.TryGetProperty("Message", out var m) ? m.GetString() : entry);
                    }
                }
                catch (JsonException)
                {
                    // non-JSON log line: ignore
                }
            }
            if (progress.Value.JobState == MigrationJobState.None) break;
            if (DateTime.UtcNow > deadline)
            {
                throw new TimeoutException($"Copy job did not finish within {_s.Copy.CopyJobTimeoutMinutes} minute(s).");
            }
            Thread.Sleep(2000);
        }
        if (errors.Count > 0)
        {
            throw new InvalidOperationException("Copy job reported: " + string.Join(" | ", errors.Distinct()));
        }
    }

    // ======================================================================= copy / move engine

    /// <summary>Copies or moves the folder and file rows of an inventory or CSV.</summary>
    private void RunContentOperation(string operationType, string modeName, IList<ContentRow> rows,
        SpConnection src, SpListInfo srcLib, SpConnection dst, SpListInfo dstLib, string destinationRootFolder)
    {
        // A same-site copy/move whose destination IS the source is stopped up front.
        if (!CheckSourceAndDestinationDiffer(src.SiteUrl, srcLib.Title, dst.SiteUrl, dstLib.Title, modeName)) return;
        if (BlockCrossTenantMove(operationType, modeName)) return;

        destinationRootFolder = (destinationRootFolder ?? "").Trim().Trim('/');
        var destinationLibraryRoot = dstLib.RootFolderUrl;
        var isCrossTenant = _opt.CrossTenant;
        var isCrossSite = !SameSite(src.SiteUrl, dst.SiteUrl);
        if (isCrossTenant)
        {
            Log("Cross-tenant operation: each file is downloaded from the source tenant and uploaded to the destination " +
                "(latest version only), then its metadata is replayed with people mapped to destination accounts.");
        }
        else if (isCrossSite)
        {
            Log("Cross-site operation detected. Files are transferred with SharePoint copy jobs.");
        }
        if (!string.IsNullOrWhiteSpace(destinationRootFolder))
        {
            EnsureFolderPath(dst, destinationLibraryRoot, destinationRootFolder);
        }

        // Folders first, shallowest first.
        var folderRows = rows.Where(r => r.ItemType == "Folder")
            .OrderBy(r => r.SourceRelativePathUnderRoot, StringComparer.OrdinalIgnoreCase).ToList();
        var folderCount = 0;
        foreach (var row in folderRows)
        {
            folderCount++;
            if (CancelRequested) break;
            _run.Progress("Ensuring destination folders", folderCount, folderRows.Count, row.SourceRelativePathUnderRoot);
            if (string.IsNullOrWhiteSpace(row.SourceRelativePathUnderRoot)) continue;
            var target = string.IsNullOrWhiteSpace(destinationRootFolder)
                ? row.SourceRelativePathUnderRoot
                : $"{destinationRootFolder}/{row.SourceRelativePathUnderRoot}";
            EnsureFolderPath(dst, destinationLibraryRoot, target);
        }

        // Writable field map once for the whole operation.
        GetWritableFieldMap(src, srcLib, dst, dstLib);

        var fileRows = rows.Where(r => r.ItemType == "File").ToList();
        var count = 0;
        foreach (var row in fileRows)
        {
            count++;
            if (CancelRequested)
            {
                Log($"Cancellation requested by operator. Stopping after {count - 1} file(s) processed.", MigratorLogLevel.Warn);
                break;
            }
            _run.Progress($"{operationType} SharePoint files", count, fileRows.Count, row.FileName);

            var timer = Stopwatch.StartNew();
            var sourceUrl = NormalizeServerRelativeUrl(row.SourceServerRelativeUrl);
            var relativePath = row.SourceRelativePathUnderRoot;
            var relativeFolder = GetParentFolderRelativePath(relativePath);
            var leafName = GetLeafName(relativePath);
            var destinationRelativeFolder = string.IsNullOrWhiteSpace(destinationRootFolder) ? relativeFolder
                : string.IsNullOrWhiteSpace(relativeFolder) ? destinationRootFolder
                : $"{destinationRootFolder}/{relativeFolder}";
            var destinationFileUrl = sourceUrl; // replaced below; keeps the catch block's log meaningful

            try
            {
                var destinationFolderUrl = EnsureFolderPath(dst, destinationLibraryRoot, destinationRelativeFolder);
                destinationFileUrl = JoinUrlPath(destinationFolderUrl, leafName);

                ListItem preloaded = null;
                try
                {
                    preloaded = src.Web.GetFileByServerRelativePath(ResourcePath.FromDecodedUrl(sourceUrl)).ListItemAllFields;
                    src.Context.Load(preloaded);
                    src.ExecuteQuery();
                }
                catch (Exception ex)
                {
                    preloaded = null;
                    _run.AddErrorLog(modeName, "PreloadSourceMetadata", sourceUrl, destinationFileUrl, ex);
                    Log($"Warning: Could not preload source metadata for {sourceUrl}. Metadata replay may be incomplete.", MigratorLogLevel.Warn);
                }

                if (_s.Copy.CheckDestinationDuplicates)
                {
                    var exists = FileExists(dst, destinationFileUrl);
                    if (exists && !_opt.OverwriteExisting && _s.Copy.SkipExistingWhenOverwriteDisabled)
                    {
                        _run.AddOperationLog(modeName, operationType, src.SiteUrl, sourceUrl, dst.SiteUrl, destinationFileUrl,
                            "File", "Skipped", "Destination file already exists and OverwriteExisting is false.", "NotApplied",
                            timer.Elapsed.TotalSeconds);
                        Log($"Skipped existing destination file: {destinationFileUrl}", MigratorLogLevel.Warn);
                        continue;
                    }
                    if (exists && _opt.OverwriteExisting)
                    {
                        Log($"Destination file exists and will be overwritten: {destinationFileUrl}", MigratorLogLevel.Warn);
                    }
                }

                if (isCrossTenant)
                {
                    TransferFileAcrossTenants(src, sourceUrl, dst, destinationFolderUrl, leafName);
                }
                else
                {
                    TransferFile(operationType, src, sourceUrl, dst, destinationFolderUrl, destinationFileUrl, isCrossSite);
                }

                var ready = !_s.Copy.WaitForDestinationFileBeforeMetadata || WaitForFile(dst, destinationFileUrl);
                string metadataStatus;
                if (ready)
                {
                    metadataStatus = ApplyFileMetadata(src, srcLib, dst, dstLib, sourceUrl, destinationFileUrl, preloaded);
                    if (isCrossTenant) metadataStatus = "LatestVersionOnly;" + metadataStatus;
                }
                else
                {
                    metadataStatus = "MetadataSkipped:DestinationFileNotFoundAfterWait";
                    _run.AddErrorLog(modeName, "WaitForDestinationFileBeforeMetadata", sourceUrl, destinationFileUrl,
                        "Destination file was not available after wait/retry period. Metadata replay skipped.",
                        "DestinationFileWaitTimeout");
                    Log($"Destination file was not available for metadata replay after wait period: {destinationFileUrl}", MigratorLogLevel.Warn);
                }

                _run.AddOperationLog(modeName, operationType, src.SiteUrl, sourceUrl, dst.SiteUrl, destinationFileUrl,
                    "File", "Success", "", metadataStatus, timer.Elapsed.TotalSeconds);
            }
            catch (Exception ex)
            {
                _run.AddOperationLog(modeName, operationType, src.SiteUrl, sourceUrl, dst.SiteUrl, destinationFileUrl,
                    "File", "Error", ex.Message, "NotApplied", timer.Elapsed.TotalSeconds);
                _run.AddErrorLog(modeName, operationType, sourceUrl, destinationFileUrl, ex);
                Log($"ERROR: {sourceUrl} -> {destinationFileUrl} :: {ex.Message}", MigratorLogLevel.Error);
            }

            if (count % 100 == 0) _run.ExportLogs(); // checkpoint, as the list copy did
        }
        Log($"{operationType} operation complete. Files processed: {count}", MigratorLogLevel.Success);
    }

    // ======================================================================= modes 2 & 3

    /// <summary>Full library/folder copy or move (GUI "FullCopy"/"FullMove").</summary>
    public void RunFullLibrary(string operationType, string sourceSiteUrl, string sourceList, string destinationSiteUrl,
        string destinationList, string destinationRootFolder)
    {
        var mode = "FullLibrary" + operationType;
        if (BlockCrossTenantMove(operationType, mode)) return;
        var src = Connect(sourceSiteUrl);
        var dst = Connect(destinationSiteUrl, TenantSide.Destination);
        var dstLib = ConfirmDestinationOrProvision(src, sourceList, dst, destinationList);
        if (dstLib == null)
        {
            Log("Destination not available. Operation aborted.", MigratorLogLevel.Warn);
            return;
        }
        var srcLib = GetListOrThrow(src, sourceList);
        var rows = InventoryLibrary(src, sourceList, $"Inventory_Internal_{RunContext.SafeFileName(sourceList)}.csv")
            .Select(r => r.ToContentRow()).ToList();
        WriteFieldCompatibilityReport(src, srcLib, dst, dstLib, mode);
        LogPreflight(mode, operationType, src.SiteUrl, sourceList, dst.SiteUrl, destinationList,
            rows.Count(r => r.ItemType == "File"), "Full library/folder operation. Folder structure will be preserved.");
        RunContentOperation(operationType, mode, rows, src, srcLib, dst, dstLib, destinationRootFolder);
    }

    // ======================================================================= modes 4 & 5

    /// <summary>CSV-driven copy or move. Site/library values come from the first CSV row, else from the window.</summary>
    public void RunCsv(string operationType, string csvPath, string sourceSiteUrl, string sourceList,
        string destinationSiteUrl, string destinationList, string destinationRootFolder)
    {
        var mode = "Csv" + operationType;
        if (BlockCrossTenantMove(operationType, mode)) return;
        if (string.IsNullOrWhiteSpace(csvPath) || !System.IO.File.Exists(csvPath))
        {
            Log($"CSV file not found: {csvPath}", MigratorLogLevel.Error);
            return;
        }
        Log($"Selected CSV: {csvPath}");
        var table = Csv.Read(csvPath);
        if (table.Rows.Count == 0)
        {
            Log("CSV contains no rows.", MigratorLogLevel.Error);
            return;
        }
        foreach (var required in RequiredCsvColumns)
        {
            if (!table.HasColumn(required)) throw new InvalidOperationException($"CSV is missing required column: {required}");
        }

        string FirstRowOr(string column, string fallback)
        {
            var first = table.Rows[0];
            return first.TryGetValue(column, out var v) && !string.IsNullOrWhiteSpace(v) ? v.Trim() : fallback;
        }

        var srcSite = FirstRowOr("SourceSiteUrl", sourceSiteUrl);
        var srcLibName = FirstRowOr("SourceLibrary", sourceList);
        var dstSite = FirstRowOr("DestinationSiteUrl", destinationSiteUrl);
        var dstLibName = FirstRowOr("DestinationLibrary", destinationList);
        var dstRoot = FirstRowOr("DestinationRootFolder", destinationRootFolder);
        if (string.IsNullOrWhiteSpace(dstSite))
        {
            if (_opt.CrossTenant)
            {
                Log("Cross-tenant CSV run: no destination site URL in the CSV's first row or on the window.", MigratorLogLevel.Error);
                return;
            }
            dstSite = srcSite;
        }
        if (string.IsNullOrWhiteSpace(dstLibName))
        {
            Log("No destination library in the CSV or on the window.", MigratorLogLevel.Error);
            return;
        }

        var rows = table.Rows.Select(r => new ContentRow
        {
            SourceSiteUrl = r.GetValueOrDefault("SourceSiteUrl"),
            SourceLibrary = r.GetValueOrDefault("SourceLibrary"),
            SourceServerRelativeUrl = r.GetValueOrDefault("SourceServerRelativeUrl"),
            SourceRelativePathUnderRoot = r.GetValueOrDefault("SourceRelativePathUnderRoot"),
            ItemType = r.GetValueOrDefault("ItemType"),
            FileName = r.GetValueOrDefault("FileName"),
            DestinationSiteUrl = r.GetValueOrDefault("DestinationSiteUrl"),
            DestinationLibrary = r.GetValueOrDefault("DestinationLibrary"),
            DestinationRootFolder = r.GetValueOrDefault("DestinationRootFolder")
        }).ToList();

        var src = Connect(srcSite);
        var dst = Connect(dstSite, TenantSide.Destination);
        var dstLib = ConfirmDestinationOrProvision(src, srcLibName, dst, dstLibName);
        if (dstLib == null)
        {
            Log("Destination not available. Operation aborted.", MigratorLogLevel.Warn);
            return;
        }
        var srcLib = GetListOrThrow(src, srcLibName);
        WriteFieldCompatibilityReport(src, srcLib, dst, dstLib, mode);
        LogPreflight(mode, operationType, src.SiteUrl, srcLibName, dst.SiteUrl, dstLibName,
            rows.Count(r => r.ItemType == "File"), "CSV-driven operation. Destination values are read from the first CSV row or the window.");
        RunContentOperation(operationType, mode, rows, src, srcLib, dst, dstLib, dstRoot);
    }

    // ======================================================================= mode 7

    /// <summary>
    /// Re-runs the failed File rows of an OperationLog.csv.
    /// Copy vs Move is taken from the log's Operation column (the v2.3.1 GUI always retried as Copy).
    /// </summary>
    public void RunRetry(string operationLogPath, string sourceList, string destinationList, string destinationRootFolder)
    {
        if (!_s.Advanced.EnableFailedItemRetryMode)
        {
            Log("Failed item retry mode is disabled in Settings.", MigratorLogLevel.Warn);
            return;
        }
        if (!System.IO.File.Exists(operationLogPath))
        {
            Log($"Operation log not found: {operationLogPath}", MigratorLogLevel.Error);
            return;
        }

        var table = Csv.Read(operationLogPath);
        var failed = table.Rows.Where(r =>
            r.GetValueOrDefault("Status") == "Error" &&
            r.GetValueOrDefault("ItemType") == "File" &&
            !string.IsNullOrWhiteSpace(r.GetValueOrDefault("SourceServerRelativeUrl"))).ToList();
        if (failed.Count == 0)
        {
            Log("No failed file rows found in selected operation log.", MigratorLogLevel.Warn);
            return;
        }
        Log($"Failed file rows found: {failed.Count}");

        var operationType = failed.Any(r => string.Equals(r.GetValueOrDefault("Operation"), "Move", StringComparison.OrdinalIgnoreCase))
            ? "Move" : "Copy";
        Log($"Retry operation type (from log): {operationType}");
        if (BlockCrossTenantMove(operationType, "RetryFailed" + operationType)) return;

        var sourceSiteUrl = failed[0].GetValueOrDefault("SourceSiteUrl");
        var destinationSiteUrl = failed[0].GetValueOrDefault("DestinationSiteUrl");
        if (string.IsNullOrWhiteSpace(destinationSiteUrl))
        {
            if (_opt.CrossTenant)
            {
                Log("Cross-tenant retry: the operation log has no DestinationSiteUrl.", MigratorLogLevel.Error);
                return;
            }
            destinationSiteUrl = sourceSiteUrl;
        }

        var src = Connect(sourceSiteUrl);
        var dst = Connect(destinationSiteUrl, TenantSide.Destination);
        var dstLib = ConfirmDestinationOrProvision(src, sourceList, dst, destinationList);
        if (dstLib == null)
        {
            Log("Destination not available. Retry aborted.", MigratorLogLevel.Warn);
            return;
        }
        var srcLib = GetListOrThrow(src, sourceList);

        var retryRows = failed.Select(r =>
        {
            var sourceUrl = NormalizeServerRelativeUrl(r.GetValueOrDefault("SourceServerRelativeUrl"));
            var relative = GetRelativePathUnderRoot(sourceUrl, srcLib.RootFolderUrl);
            return new ContentRow
            {
                SourceSiteUrl = src.SiteUrl,
                SourceLibrary = sourceList,
                SourceServerRelativeUrl = sourceUrl,
                SourceRelativePathUnderRoot = relative,
                ItemType = "File",
                FileName = GetLeafName(relative),
                DestinationSiteUrl = dst.SiteUrl,
                DestinationLibrary = destinationList,
                DestinationRootFolder = destinationRootFolder
            };
        }).ToList();

        var retryCsv = _run.ReportPath("RetryInput_FromFailedOperationLog.csv");
        Csv.Write(retryCsv, retryRows);
        Log($"Retry input CSV generated: {retryCsv}", MigratorLogLevel.Success);

        var mode = "RetryFailed" + operationType;
        WriteFieldCompatibilityReport(src, srcLib, dst, dstLib, mode);
        RunContentOperation(operationType, mode, retryRows, src, srcLib, dst, dstLib, destinationRootFolder);
    }

    // ======================================================================= mode 8

    /// <summary>CSV template -> Reports\Mode4_Mode5_CSV_Template.csv</summary>
    public void GenerateCsvTemplate()
    {
        if (!_s.Advanced.EnableCsvTemplateGenerator)
        {
            Log("CSV template generator is disabled in Settings.", MigratorLogLevel.Warn);
            return;
        }
        var tenantRoot = SpoAuthenticator.SharePointRoot(_s.Connection.TenantName) ?? "https://tenant.sharepoint.com";
        var rows = new[]
        {
            new ContentRow
            {
                SourceSiteUrl = $"{tenantRoot}/sites/SourceSite", SourceLibrary = "SourceLibrary",
                SourceServerRelativeUrl = "/sites/SourceSite/SourceLibrary/Folder/File.docx",
                SourceRelativePathUnderRoot = "Folder/File.docx", ItemType = "File", FileName = "File.docx",
                DestinationSiteUrl = $"{tenantRoot}/sites/DestinationSite", DestinationLibrary = "DestinationLibrary",
                DestinationRootFolder = "OptionalRootFolder"
            },
            new ContentRow
            {
                SourceSiteUrl = $"{tenantRoot}/sites/SourceSite", SourceLibrary = "SourceLibrary",
                SourceServerRelativeUrl = "/sites/SourceSite/SourceLibrary/Folder",
                SourceRelativePathUnderRoot = "Folder", ItemType = "Folder", FileName = "Folder",
                DestinationSiteUrl = $"{tenantRoot}/sites/DestinationSite", DestinationLibrary = "DestinationLibrary",
                DestinationRootFolder = "OptionalRootFolder"
            }
        };
        var path = _run.ReportPath("Mode4_Mode5_CSV_Template.csv");
        Csv.Write(path, rows);
        Log($"CSV template generated: {path}", MigratorLogLevel.Success);

        // Starter file for Settings > Cross-tenant > User mapping CSV.
        var sourceDomain = SampleDomain(_s.Connection.TenantName, "contoso.com");
        var destinationDomain = SampleDomain(_s.Destination.TenantName, "fabrikam.com");
        var mappingPath = _run.ReportPath("UserMapping_Template.csv");
        Csv.Write(mappingPath, new[]
        {
            new UserMappingTemplateRow { SourceUser = $"jane.doe@{sourceDomain}", DestinationUser = $"jane.doe@{destinationDomain}" },
            new UserMappingTemplateRow { SourceUser = $"john.smith@{sourceDomain}", DestinationUser = $"jsmith@{destinationDomain}" }
        });
        Log($"User mapping template generated (cross-tenant): {mappingPath}", MigratorLogLevel.Success);
    }

    private sealed class UserMappingTemplateRow
    {
        public string SourceUser { get; set; }
        public string DestinationUser { get; set; }
    }

    /// <summary>"contoso" -> "contoso.com" for sample rows; a GUID or blank tenant gives the fallback.</summary>
    private static string SampleDomain(string tenantName, string fallback)
    {
        var t = (tenantName ?? "").Trim().ToLowerInvariant();
        if (t.Length == 0 || Guid.TryParse(t, out _)) return fallback;
        if (t.EndsWith(".onmicrosoft.com")) return t;
        return t.Split('.')[0] + ".com";
    }
}
