using System.IO;
using System.Reflection;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace ContentMigrator.Config;

/// <summary>
/// Loads and saves appsettings.json next to the executable. The file is plain JSON and
/// can be edited by hand (comments and trailing commas are tolerated on load) or from the
/// Settings window. Saving from the app rewrites the file and drops any comments.
/// </summary>
public static class SettingsStore
{
    public const string FileName = "appsettings.json";

    public static string SettingsPath => Path.Combine(AppContext.BaseDirectory, FileName);

    private static readonly JsonSerializerOptions ReadOptions = new()
    {
        PropertyNameCaseInsensitive = true,
        ReadCommentHandling = JsonCommentHandling.Skip,
        AllowTrailingCommas = true
    };

    private static readonly JsonSerializerOptions WriteOptions = new()
    {
        WriteIndented = true,
        DefaultIgnoreCondition = JsonIgnoreCondition.Never
    };

    public sealed class LoadResult
    {
        public AppSettings Settings { get; init; }
        public bool CreatedNew { get; init; }
        public string Error { get; init; }
    }

    /// <summary>
    /// Loads settings. If the file does not exist, defaults are written to it. If the file
    /// cannot be parsed, defaults are returned with an error and the file is NOT overwritten,
    /// so a hand-edit mistake is never silently lost.
    /// </summary>
    public static LoadResult Load()
    {
        var path = SettingsPath;
        if (!File.Exists(path))
        {
            var defaults = new AppSettings();
            try
            {
                Save(defaults);
                return new LoadResult { Settings = defaults, CreatedNew = true };
            }
            catch (Exception ex)
            {
                return new LoadResult
                {
                    Settings = defaults,
                    CreatedNew = true,
                    Error = $"Could not create {path}: {ex.Message}"
                };
            }
        }

        try
        {
            var json = File.ReadAllText(path);
            var settings = JsonSerializer.Deserialize<AppSettings>(json, ReadOptions) ?? new AppSettings();
            FillMissingSections(settings);
            MigrateLegacyValues(json, settings);
            return new LoadResult { Settings = settings };
        }
        catch (Exception ex)
        {
            return new LoadResult
            {
                Settings = new AppSettings(),
                Error = $"{path} could not be read, so built-in defaults are in use for this session. " +
                        $"Fix the file or save from Settings to replace it.\n\n{ex.Message}"
            };
        }
    }

    /// <summary>
    /// Carries 1.0 values whose keys moved. 1.0 kept RequireSiteCollectionAdmin under Connection;
    /// it now lives under Advanced. Without this an operator's "false" would silently become true.
    /// </summary>
    private static void MigrateLegacyValues(string json, AppSettings settings)
    {
        try
        {
            using var doc = JsonDocument.Parse(json, new JsonDocumentOptions
            {
                CommentHandling = JsonCommentHandling.Skip,
                AllowTrailingCommas = true
            });
            var root = doc.RootElement;
            if (root.ValueKind != JsonValueKind.Object) return;

            var advanced = FindProperty(root, "Advanced");
            if (advanced.HasValue && FindProperty(advanced.Value, "RequireSiteCollectionAdmin").HasValue) return;

            var connection = FindProperty(root, "Connection");
            var legacy = connection.HasValue ? FindProperty(connection.Value, "RequireSiteCollectionAdmin") : null;
            if (legacy.HasValue && legacy.Value.ValueKind is JsonValueKind.True or JsonValueKind.False)
            {
                settings.Advanced.RequireSiteCollectionAdmin = legacy.Value.GetBoolean();
            }
        }
        catch (JsonException)
        {
            // The main deserialize already succeeded; a migration problem must not block loading.
        }
    }

    /// <summary>Case-insensitive property lookup, matching how the settings are read.</summary>
    private static JsonElement? FindProperty(JsonElement obj, string name)
    {
        if (obj.ValueKind != JsonValueKind.Object) return null;
        foreach (var p in obj.EnumerateObject())
        {
            if (string.Equals(p.Name, name, StringComparison.OrdinalIgnoreCase)) return p.Value;
        }
        return null;
    }

    /// <summary>Writes atomically (temp file, then replace) so a crash never leaves half a file.</summary>
    public static void Save(AppSettings settings)
    {
        var path = SettingsPath;
        var temp = path + ".tmp";
        var json = JsonSerializer.Serialize(settings, WriteOptions);
        File.WriteAllText(temp, json);
        if (File.Exists(path))
        {
            File.Replace(temp, path, null);
        }
        else
        {
            File.Move(temp, path);
        }
    }

    public static AppSettings Clone(AppSettings settings)
    {
        var json = JsonSerializer.Serialize(settings, WriteOptions);
        var copy = JsonSerializer.Deserialize<AppSettings>(json, ReadOptions);
        FillMissingSections(copy);
        return copy;
    }

    /// <summary>A section removed from a hand-edited file comes back with defaults instead of null.</summary>
    private static void FillMissingSections(AppSettings settings)
    {
        foreach (var prop in typeof(AppSettings).GetProperties(BindingFlags.Public | BindingFlags.Instance))
        {
            if (prop.GetValue(settings) == null)
            {
                prop.SetValue(settings, Activator.CreateInstance(prop.PropertyType));
            }
            var section = prop.GetValue(settings);
            foreach (var inner in prop.PropertyType.GetProperties(BindingFlags.Public | BindingFlags.Instance))
            {
                if (inner.PropertyType == typeof(List<string>) && inner.GetValue(section) == null)
                {
                    inner.SetValue(section, new List<string>());
                }
                if (inner.PropertyType == typeof(string) && inner.GetValue(section) == null)
                {
                    inner.SetValue(section, "");
                }
            }
        }
    }

    /// <summary>Extra checks for a cross-tenant run; empty when the destination tenant is usable.</summary>
    public static List<string> ValidateCrossTenant(AppSettings settings)
    {
        var problems = new List<string>();
        if (string.IsNullOrWhiteSpace(settings.Destination.TenantName))
        {
            problems.Add("Destination tenant > Tenant name is required for a cross-tenant run.");
        }
        var csv = settings.CrossTenant.UserMappingCsvPath;
        if (!string.IsNullOrWhiteSpace(csv) && !File.Exists(csv))
        {
            problems.Add($"Cross-tenant > User mapping CSV was not found: {csv}");
        }
        return problems;
    }

    /// <summary>"contoso.com=fabrikam.com" (spaces and a leading @ tolerated).</summary>
    public static bool TryParseDomainMapping(string line, out string from, out string to)
    {
        from = to = null;
        var parts = (line ?? "").Split('=');
        if (parts.Length != 2) return false;
        from = parts[0].Trim().TrimStart('@');
        to = parts[1].Trim().TrimStart('@');
        return from.Length > 0 && to.Length > 0 && !from.Contains(' ') && !to.Contains(' ');
    }

    /// <summary>Returns human-readable problems; empty when the settings are usable.</summary>
    public static List<string> Validate(AppSettings settings)
    {
        var problems = new List<string>();
        foreach (var sectionProp in typeof(AppSettings).GetProperties(BindingFlags.Public | BindingFlags.Instance))
        {
            var sectionTitle = sectionProp.GetCustomAttribute<SettingsSectionAttribute>()?.Title ?? sectionProp.Name;
            var section = sectionProp.GetValue(settings);
            foreach (var prop in sectionProp.PropertyType.GetProperties(BindingFlags.Public | BindingFlags.Instance))
            {
                var meta = prop.GetCustomAttribute<SettingAttribute>();
                if (meta == null) continue;
                var value = prop.GetValue(section);
                if (meta.Required && prop.PropertyType == typeof(string) && string.IsNullOrWhiteSpace((string)value))
                {
                    problems.Add($"{sectionTitle} > {meta.Label} is required.");
                }
                if (prop.PropertyType == typeof(int))
                {
                    var i = (int)value;
                    if (i < meta.Min || i > meta.Max)
                    {
                        problems.Add($"{sectionTitle} > {meta.Label} must be between {meta.Min} and {meta.Max}.");
                    }
                }
            }
        }

        if (!string.IsNullOrWhiteSpace(settings.Connection.ClientId) && !Guid.TryParse(settings.Connection.ClientId, out _))
        {
            problems.Add("Connection > App (client) ID must be a GUID.");
        }
        if (!string.IsNullOrWhiteSpace(settings.Connection.RedirectUri) &&
            !Uri.TryCreate(settings.Connection.RedirectUri, UriKind.Absolute, out _))
        {
            problems.Add("Connection > Redirect URI must be an absolute URI such as http://localhost.");
        }
        var dest = settings.Destination;
        if (!string.IsNullOrWhiteSpace(dest.ClientId) && !Guid.TryParse(dest.ClientId, out _))
        {
            problems.Add("Destination tenant > App (client) ID must be a GUID (or blank to use the source App ID).");
        }
        if (!string.IsNullOrWhiteSpace(dest.RedirectUri) && !Uri.TryCreate(dest.RedirectUri, UriKind.Absolute, out _))
        {
            problems.Add("Destination tenant > Redirect URI must be an absolute URI such as http://localhost (or blank).");
        }
        foreach (var line in settings.CrossTenant.DomainMappings ?? new List<string>())
        {
            if (!TryParseDomainMapping(line, out _, out _))
            {
                problems.Add($"Cross-tenant > Domain mappings: '{line}' must look like contoso.com=fabrikam.com.");
            }
        }
        try
        {
            _ = new System.Text.RegularExpressions.Regex(settings.Naming.UrlIllegalCharacterPattern ?? "");
        }
        catch (ArgumentException ex)
        {
            problems.Add($"Naming > Illegal URL character pattern is not a valid regular expression: {ex.Message}");
        }
        return problems;
    }
}
