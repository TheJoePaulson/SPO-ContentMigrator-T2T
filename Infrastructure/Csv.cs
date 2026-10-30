using System.IO;
using System.Globalization;
using System.Reflection;
using System.Text;

namespace ContentMigrator.Infrastructure;

/// <summary>
/// Minimal RFC 4180 CSV reader/writer (no third-party package). Output matches what
/// Export-Csv -NoTypeInformation produced: a header row, every field double-quoted,
/// UTF-8. A BOM is written so Excel opens non-ASCII names correctly.
/// </summary>
public static class Csv
{
    private static readonly UTF8Encoding Utf8Bom = new(encoderShouldEmitUTF8Identifier: true);

    /// <summary>Public instance properties in declaration order become the columns.</summary>
    public static void Write<T>(string path, IEnumerable<T> rows)
    {
        var props = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance)
            .Where(p => p.CanRead)
            .OrderBy(p => p.MetadataToken)
            .ToArray();

        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        using var writer = new StreamWriter(path, false, Utf8Bom);
        writer.WriteLine(string.Join(",", props.Select(p => Quote(p.Name))));
        foreach (var row in rows)
        {
            writer.WriteLine(string.Join(",", props.Select(p => Quote(Format(p.GetValue(row))))));
        }
    }

    public static string Format(object value)
    {
        return value switch
        {
            null => "",
            string s => s,
            DateTime dt => dt.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture),
            DateTimeOffset dto => dto.ToString("yyyy-MM-dd HH:mm:ss zzz", CultureInfo.InvariantCulture),
            bool b => b ? "True" : "False",
            IFormattable f => f.ToString(null, CultureInfo.InvariantCulture),
            _ => value.ToString()
        };
    }

    private static string Quote(string value) => "\"" + (value ?? "").Replace("\"", "\"\"") + "\"";

    public sealed class Table
    {
        public List<string> Headers { get; } = new();
        public List<Dictionary<string, string>> Rows { get; } = new();

        public bool HasColumn(string name) => Headers.Any(h => string.Equals(h, name, StringComparison.OrdinalIgnoreCase));
    }

    /// <summary>Reads a CSV with a header row. Keys are case-insensitive; missing cells read as "".</summary>
    public static Table Read(string path)
    {
        var text = File.ReadAllText(path, Encoding.UTF8);
        var records = Parse(text);
        var table = new Table();
        if (records.Count == 0) return table;

        table.Headers.AddRange(records[0].Select(h => h.Trim()));
        for (var r = 1; r < records.Count; r++)
        {
            var record = records[r];
            if (record.Count == 1 && string.IsNullOrEmpty(record[0])) continue; // blank line
            var row = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            for (var c = 0; c < table.Headers.Count; c++)
            {
                row[table.Headers[c]] = c < record.Count ? record[c] : "";
            }
            table.Rows.Add(row);
        }
        return table;
    }

    private static List<List<string>> Parse(string text)
    {
        var records = new List<List<string>>();
        var record = new List<string>();
        var field = new StringBuilder();
        var inQuotes = false;
        var i = 0;
        if (text.Length > 0 && text[0] == '\uFEFF') i = 1;

        for (; i < text.Length; i++)
        {
            var ch = text[i];
            if (inQuotes)
            {
                if (ch == '"')
                {
                    if (i + 1 < text.Length && text[i + 1] == '"') { field.Append('"'); i++; }
                    else inQuotes = false;
                }
                else field.Append(ch);
                continue;
            }

            switch (ch)
            {
                case '"':
                    inQuotes = true;
                    break;
                case ',':
                    record.Add(field.ToString());
                    field.Clear();
                    break;
                case '\r':
                    break;
                case '\n':
                    record.Add(field.ToString());
                    field.Clear();
                    records.Add(record);
                    record = new List<string>();
                    break;
                default:
                    field.Append(ch);
                    break;
            }
        }
        if (field.Length > 0 || record.Count > 0)
        {
            record.Add(field.ToString());
            records.Add(record);
        }
        return records;
    }
}
