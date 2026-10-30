using ContentMigrator.Config;
using ContentMigrator.Infrastructure;
using ContentMigrator.SharePoint;

namespace ContentMigrator.Engine;

public enum OperationKind
{
    Inventory,
    FullMove,
    FullCopy,
    CsvMove,
    CsvCopy,
    ListItemCopy,
    Retry,
    CsvTemplate,
    CreateFromExisting,
    ModernizeList
}

/// <summary>Everything the main window collects for one run.</summary>
public sealed class OperationRequest
{
    public OperationKind Operation { get; init; }
    public string SourceSiteUrl { get; init; } = "";
    public string SourceList { get; init; } = "";
    public string DestinationSiteUrl { get; init; } = "";
    public string DestinationList { get; init; } = "";
    public string DestinationRootFolder { get; init; } = "";
    public string CsvPath { get; init; } = "";
    public RunOptions Options { get; init; } = new();
}

/// <summary>
/// Builds the run context, runs one operation on a background
/// thread, always exports the logs, and reports the run folders back to the window.
/// </summary>
public static class OperationRunner
{
    /// <param name="destinationAuth">Destination tenant sign-in; only used when request.Options.CrossTenant is set.</param>
    public static void Run(AppSettings settings, SpoAuthenticator sourceAuth, SpoAuthenticator destinationAuth,
        OperationRequest request, RunContext run, CancellationToken cancellationToken)
    {
        run.CrossTenant = request.Options.CrossTenant;
        run.SourceTenant = settings.Connection.TenantName;
        run.DestinationTenant = request.Options.CrossTenant ? settings.Destination.TenantName : settings.Connection.TenantName;
        using var engine = new MigrationEngine(settings, run, sourceAuth, destinationAuth, request.Options, cancellationToken);
        try
        {
            run.InitializeFolders();
            engine.CheckPrerequisites();
            run.Log($"Starting operation: {request.Operation}", MigratorLogLevel.Success);

            var r = request;
            switch (r.Operation)
            {
                case OperationKind.Inventory:
                    engine.InventoryLibrary(engine.Connect(r.SourceSiteUrl), r.SourceList);
                    break;
                case OperationKind.FullMove:
                    engine.RunFullLibrary("Move", r.SourceSiteUrl, r.SourceList, r.DestinationSiteUrl, r.DestinationList, r.DestinationRootFolder);
                    break;
                case OperationKind.FullCopy:
                    engine.RunFullLibrary("Copy", r.SourceSiteUrl, r.SourceList, r.DestinationSiteUrl, r.DestinationList, r.DestinationRootFolder);
                    break;
                case OperationKind.CsvMove:
                    engine.RunCsv("Move", r.CsvPath, r.SourceSiteUrl, r.SourceList, r.DestinationSiteUrl, r.DestinationList, r.DestinationRootFolder);
                    break;
                case OperationKind.CsvCopy:
                    engine.RunCsv("Copy", r.CsvPath, r.SourceSiteUrl, r.SourceList, r.DestinationSiteUrl, r.DestinationList, r.DestinationRootFolder);
                    break;
                case OperationKind.ListItemCopy:
                    engine.RunListItemCopy(r.SourceSiteUrl, r.SourceList, r.DestinationSiteUrl, r.DestinationList);
                    break;
                case OperationKind.Retry:
                    engine.RunRetry(r.CsvPath, r.SourceList, r.DestinationList, r.DestinationRootFolder);
                    break;
                case OperationKind.CsvTemplate:
                    engine.GenerateCsvTemplate();
                    break;
                case OperationKind.CreateFromExisting:
                    engine.RunCreateFromExisting(r.SourceSiteUrl, r.SourceList, r.DestinationSiteUrl, r.DestinationList);
                    break;
                case OperationKind.ModernizeList:
                    engine.RunModernize(r.SourceSiteUrl, r.SourceList, r.DestinationSiteUrl, r.DestinationList);
                    break;
                default:
                    run.Log($"Unknown operation: {r.Operation}", MigratorLogLevel.Error);
                    break;
            }

            if (cancellationToken.IsCancellationRequested)
            {
                run.Log("Operation cancelled by operator. Work completed before the cancel is kept and logged.", MigratorLogLevel.Warn);
            }
            run.Complete();
            run.Log($"Operation finished. Output folder: {run.RunFolder}", MigratorLogLevel.Success);
        }
        catch (SiteAdminRequiredException ex)
        {
            run.BlockingMessage = ex.Message; // shown as a dialog by the window
            try { run.Complete(); } catch (Exception exportEx) { run.Log($"Log export failed: {exportEx.Message}", MigratorLogLevel.Error); }
        }
        catch (Exception ex)
        {
            run.Log($"Operation failed: {ex.Message}", MigratorLogLevel.Error);
            run.AddErrorLog("GUI", request.Operation.ToString(), "", "", ex);
            try { run.Complete(); } catch (Exception exportEx) { run.Log($"Log export failed: {exportEx.Message}", MigratorLogLevel.Error); }
        }
    }
}
