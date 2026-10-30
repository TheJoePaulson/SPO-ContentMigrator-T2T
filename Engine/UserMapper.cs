using System.IO;
using ContentMigrator.Config;
using ContentMigrator.Infrastructure;

namespace ContentMigrator.Engine;

/// <summary>
/// Matches people in the source tenant to accounts in the destination tenant for a
/// cross-tenant run. Order: the user mapping CSV, then the domain mappings, then the identity
/// unchanged (useful when both tenants share a verified domain alias). The fallback account
/// is applied by the caller only when the mapped account cannot be resolved on the destination.
/// </summary>
public sealed class UserMapper
{
    private readonly CrossTenantSettings _settings;
    private readonly RunContext _run;
    private readonly Dictionary<string, string> _map = new(StringComparer.OrdinalIgnoreCase);
    private readonly List<(string From, string To)> _domains = new();
    private bool _loaded;

    public UserMapper(CrossTenantSettings settings, RunContext run)
    {
        _settings = settings;
        _run = run;
    }

    public string FallbackUser => (_settings.FallbackUser ?? "").Trim();

    /// <summary>Reads the mapping CSV and domain mappings once, logging what was loaded.</summary>
    public void Load()
    {
        if (_loaded) return;
        _loaded = true;

        var path = (_settings.UserMappingCsvPath ?? "").Trim();
        if (path.Length == 0)
        {
            _run.Log("User mapping CSV: none configured.");
        }
        else if (!File.Exists(path))
        {
            _run.Log($"User mapping CSV not found: {path}. Continuing with domain mappings only.", MigratorLogLevel.Warn);
        }
        else
        {
            var table = Csv.Read(path);
            if (!table.HasColumn("SourceUser") || !table.HasColumn("DestinationUser"))
            {
                throw new InvalidOperationException(
                    $"User mapping CSV {path} must have the columns SourceUser and DestinationUser.");
            }
            var duplicates = 0;
            foreach (var row in table.Rows)
            {
                var from = Normalize(row.GetValueOrDefault("SourceUser"));
                var to = Normalize(row.GetValueOrDefault("DestinationUser"));
                if (from.Length == 0 || to.Length == 0) continue;
                if (_map.ContainsKey(from)) duplicates++;
                _map[from] = to;
            }
            _run.Log($"User mapping CSV: {_map.Count} mapping(s) loaded from {path}.", MigratorLogLevel.Success);
            if (duplicates > 0)
            {
                _run.Log($"User mapping CSV: {duplicates} duplicate SourceUser row(s); the last one wins.", MigratorLogLevel.Warn);
            }
        }

        foreach (var line in _settings.DomainMappings ?? new List<string>())
        {
            if (SettingsStore.TryParseDomainMapping(line, out var from, out var to)) _domains.Add((from, to));
        }
        _run.Log(_domains.Count == 0
            ? "Domain mappings: none configured."
            : "Domain mappings: " + string.Join(", ", _domains.Select(d => $"{d.From} -> {d.To}")) + ".");
        _run.Log(FallbackUser.Length == 0
            ? "Fallback account: none (unmatched people are left blank and logged)."
            : $"Fallback account: {FallbackUser}.");
    }

    /// <summary>The destination identity to try for a source person, and how it was found.</summary>
    public (string Candidate, string Method) Map(string sourceIdentity)
    {
        Load();
        var id = Normalize(sourceIdentity);
        if (id.Length == 0) return ("", "Empty");

        if (_map.TryGetValue(id, out var mapped)) return (mapped, "MappingCsv");

        var at = id.LastIndexOf('@');
        if (at > 0)
        {
            var domain = id.Substring(at + 1);
            foreach (var (from, to) in _domains)
            {
                if (string.Equals(domain, from, StringComparison.OrdinalIgnoreCase))
                {
                    return (id.Substring(0, at + 1) + to, "DomainSwap");
                }
            }
        }
        return (id, "Unchanged");
    }

    /// <summary>Strips a claims prefix ("i:0#.f|membership|jane@contoso.com" -> "jane@contoso.com").</summary>
    private static string Normalize(string value)
    {
        var v = (value ?? "").Trim();
        var bar = v.LastIndexOf('|');
        return bar >= 0 ? v.Substring(bar + 1).Trim() : v;
    }
}
