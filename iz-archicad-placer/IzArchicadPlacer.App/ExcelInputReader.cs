using ClosedXML.Excel;
using ExcelDataReader;

namespace IzArchicadPlacer.App;

internal static class ExcelInputReader
{
    private static readonly string[] RequiredColumns = ["filename", "minx", "miny", "text1"];

    public static List<SheetData> ReadAll(string path)
    {
        string ext = Path.GetExtension(path).ToLowerInvariant();
        return ext == ".xls" ? ReadXls(path) : ReadXlsx(path);
    }

    private static List<SheetData> ReadXlsx(string path)
    {
        using var wb = new XLWorkbook(path);
        var result = new List<SheetData>();
        foreach (var ws in wb.Worksheets)
        {
            var headers = ReadHeaders(ws.Row(1).CellsUsed().ToDictionary(c => c.Address.ColumnNumber, c => c.GetString()));
            if (!HasRequired(headers))
            {
                continue;
            }

            var rows = new List<SourceRow>();
            foreach (var row in ws.RowsUsed().Skip(1))
            {
                string filename = GetString(row, headers, "filename");
                if (string.IsNullOrWhiteSpace(filename))
                {
                    continue;
                }

                rows.Add(new SourceRow
                {
                    Filename = filename,
                    MinX = GetNumber(row, headers, "minx"),
                    MinY = GetNumber(row, headers, "miny"),
                    Text1 = GetString(row, headers, "text1"),
                    Text2 = GetString(row, headers, "text2")
                });
            }

            result.Add(new SheetData { SheetName = ws.Name, Rows = rows });
        }

        return result;
    }

    private static List<SheetData> ReadXls(string path)
    {
        var result = new List<SheetData>();
        using var stream = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        using var reader = ExcelReaderFactory.CreateReader(stream);

        do
        {
            string sheetName = reader.Name;
            var rowsRaw = new List<object?[]>();
            while (reader.Read())
            {
                var arr = new object?[reader.FieldCount];
                for (int i = 0; i < reader.FieldCount; i++)
                {
                    arr[i] = reader.GetValue(i);
                }

                rowsRaw.Add(arr);
            }

            if (rowsRaw.Count == 0)
            {
                continue;
            }

            var hdrRaw = new Dictionary<int, string>();
            for (int i = 0; i < rowsRaw[0].Length; i++)
            {
                hdrRaw[i + 1] = rowsRaw[0][i]?.ToString() ?? string.Empty;
            }

            var headers = ReadHeaders(hdrRaw);
            if (!HasRequired(headers))
            {
                continue;
            }

            var rows = new List<SourceRow>();
            foreach (var row in rowsRaw.Skip(1))
            {
                string filename = GetString(row, headers, "filename");
                if (string.IsNullOrWhiteSpace(filename))
                {
                    continue;
                }

                rows.Add(new SourceRow
                {
                    Filename = filename,
                    MinX = GetNumber(row, headers, "minx"),
                    MinY = GetNumber(row, headers, "miny"),
                    Text1 = GetString(row, headers, "text1"),
                    Text2 = GetString(row, headers, "text2")
                });
            }

            result.Add(new SheetData { SheetName = sheetName, Rows = rows });
        } while (reader.NextResult());

        return result;
    }

    private static Dictionary<string, int> ReadHeaders(Dictionary<int, string> raw)
    {
        var map = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        foreach (var pair in raw)
        {
            string key = pair.Value.Trim().ToLowerInvariant();
            if (!string.IsNullOrWhiteSpace(key))
            {
                map[key] = pair.Key;
            }
        }

        return map;
    }

    private static bool HasRequired(Dictionary<string, int> headers)
    {
        return RequiredColumns.All(headers.ContainsKey);
    }

    private static string GetString(IXLRow row, IReadOnlyDictionary<string, int> headers, string col)
    {
        if (!headers.TryGetValue(col, out int index))
        {
            return string.Empty;
        }

        return row.Cell(index).GetString();
    }

    private static double GetNumber(IXLRow row, IReadOnlyDictionary<string, int> headers, string col)
    {
        if (!headers.TryGetValue(col, out int index))
        {
            return 0d;
        }

        return row.Cell(index).TryGetValue<double>(out double v) ? v : 0d;
    }

    private static string GetString(object?[] row, IReadOnlyDictionary<string, int> headers, string col)
    {
        if (!headers.TryGetValue(col, out int oneBased))
        {
            return string.Empty;
        }

        int idx = oneBased - 1;
        if (idx < 0 || idx >= row.Length)
        {
            return string.Empty;
        }

        return row[idx]?.ToString() ?? string.Empty;
    }

    private static double GetNumber(object?[] row, IReadOnlyDictionary<string, int> headers, string col)
    {
        string s = GetString(row, headers, col);
        return double.TryParse(s, out double v) ? v : 0d;
    }
}
