using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using Microsoft.SharePoint.Client;
using ContentMigrator.Infrastructure;
using ContentMigrator.SharePoint;
using IOFile = System.IO.File;
using IODirectory = System.IO.Directory;
using IOPath = System.IO.Path;
using IOFileMode = System.IO.FileMode;
using IOFileStream = System.IO.FileStream;
using IOFileAccess = System.IO.FileAccess;
using IOFileShare = System.IO.FileShare;
using IOMemoryStream = System.IO.MemoryStream;
using IOStream = System.IO.Stream;

namespace ContentMigrator.Engine;

/// <summary>
/// Cross-tenant file transfer. SharePoint copy jobs and MoveCopyUtil only work inside one
/// tenant, so each file is streamed down from the source tenant to a staging file and uploaded
/// to the destination tenant (chunked above the configured size). Latest version only.
/// </summary>
public sealed partial class MigrationEngine
{
    /// <summary>One client for the whole app: downloads stream, so no overall timeout.</summary>
    private static readonly HttpClient Http = new(new SocketsHttpHandler
    {
        PooledConnectionLifetime = TimeSpan.FromMinutes(5),
        AutomaticDecompression = DecompressionMethods.None
    })
    {
        Timeout = Timeout.InfiniteTimeSpan
    };

    /// <summary>Copies one file from the source tenant into a destination folder. Returns the bytes copied.</summary>
    private long TransferFileAcrossTenants(SpConnection src, string sourceUrl, SpConnection dst,
        string destinationFolderUrl, string fileName)
    {
        var stagingRoot = _run.RunPath("TransferStaging");
        IODirectory.CreateDirectory(stagingRoot);
        // A GUID name avoids any character SharePoint allows but Windows does not.
        var localPath = IOPath.Combine(stagingRoot, Guid.NewGuid().ToString("N") + IOPath.GetExtension(fileName));

        var destinationFileUrl = JoinUrlPath(destinationFolderUrl, fileName);
        try
        {
            var downloaded = DownloadToFile(src, sourceUrl, localPath);
            var existedBefore = FileExists(dst, destinationFileUrl);
            UploadFromFile(dst, destinationFolderUrl, fileName, localPath, existedBefore);

            // Verify the upload. SharePoint rewrites document properties inside Office files
            // (.docx, .xlsx, ...) on upload, so their stored size can legitimately differ: a
            // mismatch is a warning. Only an empty result for a non-empty file is a failure.
            var uploaded = dst.Web.GetFileByServerRelativePath(ResourcePath.FromDecodedUrl(destinationFileUrl));
            dst.Context.Load(uploaded, f => f.Length);
            dst.ExecuteQuery();
            if (uploaded.Length == 0 && downloaded > 0)
            {
                RemoveIncompleteUpload(dst, destinationFileUrl, existedBefore);
                throw new InvalidOperationException(
                    $"Upload produced an empty file (downloaded {downloaded:N0} bytes).");
            }
            if (uploaded.Length != downloaded)
            {
                var note = $"Size differs after upload: downloaded {downloaded:N0} bytes, destination has {uploaded.Length:N0}. " +
                           "Expected for Office documents (SharePoint updates their embedded properties); check other file types.";
                Log($"  {fileName}: {note}", MigratorLogLevel.Debug);
                _run.AddErrorLog("CrossTenantCopy", "SizeCheck", sourceUrl, destinationFileUrl, note, "SizeMismatchWarning");
            }
            return downloaded;
        }
        finally
        {
            if (!_s.CrossTenant.KeepStagingFiles)
            {
                try { if (IOFile.Exists(localPath)) IOFile.Delete(localPath); } catch { /* best effort */ }
            }
        }
    }

    /// <summary>
    /// Streams a file to disk with the REST API (GetFileByServerRelativePath/$value), retrying on
    /// HTTP 429/503 and honouring Retry-After. Returns the number of bytes written.
    /// </summary>
    private long DownloadToFile(SpConnection src, string serverRelativeUrl, string localPath)
    {
        // The path goes in a parameter alias so slashes and quotes are encoded safely.
        var escaped = Uri.EscapeDataString(serverRelativeUrl.Replace("'", "''"));
        var url = $"{src.SiteUrl}/_api/web/GetFileByServerRelativePath(decodedurl=@p)/$value?@p='{escaped}'";
        var siteUri = new Uri(src.SiteUrl);

        var attempt = 0;
        var delaySeconds = _s.Advanced.RetryBaseDelaySeconds;
        while (true)
        {
            _ct.ThrowIfCancellationRequested();
            using var request = new HttpRequestMessage(HttpMethod.Get, url);
            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", src.Auth.GetAccessToken(siteUri));
            request.Headers.UserAgent.ParseAdd("ContentMigrator-T2T-CE/1.1");
            using var response = Http.Send(request, HttpCompletionOption.ResponseHeadersRead, _ct);

            var status = (int)response.StatusCode;
            if ((status == 429 || status == 503) && attempt < _s.Advanced.MaxRetryAttempts)
            {
                attempt++;
                var wait = response.Headers.RetryAfter?.Delta?.TotalSeconds ?? delaySeconds;
                Log($"Source tenant is throttling downloads (HTTP {status}). Waiting {wait:0} s (retry {attempt}).", MigratorLogLevel.Warn);
                Thread.Sleep(TimeSpan.FromSeconds(Math.Max(1, wait)));
                delaySeconds = Math.Min(delaySeconds * 2, 300);
                continue;
            }
            if (!response.IsSuccessStatusCode)
            {
                throw new InvalidOperationException($"Download from the source tenant failed: HTTP {status} {response.ReasonPhrase}.");
            }

            using var body = response.Content.ReadAsStream(_ct);
            using var file = new IOFileStream(localPath, IOFileMode.Create, IOFileAccess.Write, IOFileShare.None, 1024 * 1024);
            body.CopyTo(file, 1024 * 1024);
            return file.Length;
        }
    }

    /// <summary>
    /// Uploads a staged file into a destination folder. Up to the chunk size: one request.
    /// Larger: StartUpload / ContinueUpload / FinishUpload, so any size works within
    /// SharePoint's own limits. Overwrite follows the main window's check box.
    /// </summary>
    private void UploadFromFile(SpConnection dst, string folderServerRelativeUrl, string fileName, string localPath,
        bool existedBefore)
    {
        var chunkBytes = Math.Max(1, _s.CrossTenant.UploadChunkSizeMB) * 1024 * 1024;
        var folder = dst.Web.GetFolderByServerRelativePath(ResourcePath.FromDecodedUrl(folderServerRelativeUrl));
        var fileUrl = JoinUrlPath(folderServerRelativeUrl, fileName);

        using var source = IOFile.OpenRead(localPath);
        var length = source.Length;

        // AddUsingPath (not Files.Add + FileCreationInformation.Url) so '#' and '%' in names work.
        var addParameters = new FileCollectionAddParameters { Overwrite = _opt.OverwriteExisting };
        if (length <= chunkBytes)
        {
            folder.Files.AddUsingPath(ResourcePath.FromDecodedUrl(fileName), addParameters, source);
            dst.ExecuteQuery();
            return;
        }

        var uploadId = Guid.NewGuid();
        var buffer = new byte[chunkBytes];
        long offset = 0;
        var first = true;
        try
        {
            while (true)
            {
                _ct.ThrowIfCancellationRequested();
                var read = ReadFull(source, buffer);
                var last = source.Position >= length;
                using var slice = new IOMemoryStream(buffer, 0, read, false);

                if (first)
                {
                    // Create the file empty, then start the upload session on it.
                    var created = folder.Files.AddUsingPath(ResourcePath.FromDecodedUrl(fileName), addParameters,
                        new IOMemoryStream());
                    var started = created.StartUpload(uploadId, slice);
                    dst.ExecuteQuery();
                    offset = started.Value;
                    first = false;
                    continue; // length > chunk size, so the first slice is never the last
                }

                var target = dst.Web.GetFileByServerRelativePath(ResourcePath.FromDecodedUrl(fileUrl));
                if (last)
                {
                    target.FinishUpload(uploadId, offset, slice);
                    dst.ExecuteQuery();
                    return;
                }
                var continued = target.ContinueUpload(uploadId, offset, slice);
                dst.ExecuteQuery();
                offset = continued.Value;
            }
        }
        catch
        {
            if (!first)
            {
                try
                {
                    dst.Web.GetFileByServerRelativePath(ResourcePath.FromDecodedUrl(fileUrl)).CancelUpload(uploadId);
                    dst.ExecuteQuery();
                }
                catch { /* the original error matters more */ }
                RemoveIncompleteUpload(dst, fileUrl, existedBefore);
            }
            throw;
        }
    }

    /// <summary>
    /// After a failed upload: a file this upload created is sent to the destination recycle bin so
    /// no empty file is left behind. A file that already existed (Overwrite on) is kept, and the log
    /// says to restore its previous version from version history.
    /// </summary>
    private void RemoveIncompleteUpload(SpConnection dst, string fileUrl, bool existedBefore)
    {
        if (existedBefore)
        {
            Log($"  Upload failed after overwriting {fileUrl}. Restore its previous version from version history.",
                MigratorLogLevel.Warn);
            return;
        }
        try
        {
            dst.Web.GetFileByServerRelativePath(ResourcePath.FromDecodedUrl(fileUrl)).Recycle();
            dst.ExecuteQuery();
            Log($"  Removed the incomplete upload {fileUrl} (sent to the destination recycle bin).", MigratorLogLevel.Debug);
        }
        catch (Exception ex)
        {
            Log($"  Could not remove the incomplete upload {fileUrl}: {ex.Message}. Delete it before retrying.", MigratorLogLevel.Warn);
        }
    }

    private static int ReadFull(IOStream stream, byte[] buffer)
    {
        var total = 0;
        while (total < buffer.Length)
        {
            var n = stream.Read(buffer, total, buffer.Length - total);
            if (n == 0) break;
            total += n;
        }
        return total;
    }
}
