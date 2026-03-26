using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using ClosedXML.Excel;
using ExcelDataReader;

namespace IzMatcher.App;

internal static class Program
{
    [STAThread]
    private static void Main()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        ApplicationConfiguration.Initialize();
        Application.Run(new MainForm());
    }
}

internal sealed class MainForm : Form
{
    private readonly TextBox _excelPath = new() { Width = 360 };
    private readonly TextBox _outputPath = new() { Width = 360 };
    private readonly TextBox _entityName = new() { Width = 160, Text = "Sheet1" };
    private readonly CheckBox _appendExisting = new() { Text = "Append to existing workbook" };

    private readonly TextBox _blockName = new() { Width = 180 };
    private readonly TextBox _blockLayer = new() { Width = 180 };
    private readonly CheckBox _blockNameRegex = new() { Text = "regex" };
    private readonly CheckBox _blockLayerRegex = new() { Text = "regex" };

    private readonly TextBox _textContent = new() { Width = 180 };
    private readonly TextBox _textLayer = new() { Width = 180 };
    private readonly TextBox _textColor = new() { Width = 180 };
    private readonly TextBox _heightMin = new() { Width = 80 };
    private readonly TextBox _heightMax = new() { Width = 80 };
    private readonly CheckBox _textContentRegex = new() { Text = "regex" };
    private readonly CheckBox _textLayerRegex = new() { Text = "regex" };
    private readonly CheckBox _textColorRegex = new() { Text = "regex" };

    private readonly TextBox _t1xMin = new() { Width = 60, Text = "0" };
    private readonly TextBox _t1yMin = new() { Width = 60, Text = "0" };
    private readonly TextBox _t1xMax = new() { Width = 60, Text = "0" };
    private readonly TextBox _t1yMax = new() { Width = 60, Text = "0" };

    private readonly TextBox _t2xMin = new() { Width = 60, Text = "0" };
    private readonly TextBox _t2yMin = new() { Width = 60, Text = "0" };
    private readonly TextBox _t2xMax = new() { Width = 60, Text = "0" };
    private readonly TextBox _t2yMax = new() { Width = 60, Text = "0" };

    private readonly Label _blockInfo = new() { AutoSize = true, Text = "Blocks: 0 / 0" };
    private readonly Label _textInfo = new() { AutoSize = true, Text = "Texts: 0 / 0" };

    private readonly DataGridView _grid = new()
    {
        Dock = DockStyle.Fill,
        ReadOnly = true,
        AllowUserToAddRows = false,
        AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
    };

    private readonly List<MatchRow> _matches = [];
    private readonly string _settingsPath;

    public MainForm()
    {
        Text = "IZ Matcher";
        Width = 1300;
        Height = 800;
        StartPosition = FormStartPosition.CenterScreen;
        TrySetWindowIcon();

        string appDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "IZBLOCK");
        Directory.CreateDirectory(appDir);
        _settingsPath = Path.Combine(appDir, "izmatcher.settings.json");

        var split = new SplitContainer
        {
            Dock = DockStyle.Fill,
            SplitterDistance = 470
        };
        Controls.Add(split);

        var left = new Panel { Dock = DockStyle.Fill, AutoScroll = true };
        split.Panel1.Controls.Add(left);

        var root = new TableLayoutPanel
        {
            Dock = DockStyle.Top,
            AutoSize = true,
            ColumnCount = 1,
            Padding = new Padding(8)
        };
        left.Controls.Add(root);

        root.Controls.Add(BuildFileGroup());
        root.Controls.Add(BuildBlockGroup());
        root.Controls.Add(BuildTextGroup());
        root.Controls.Add(BuildCoordGroup("Text1 Window", _t1xMin, _t1yMin, _t1xMax, _t1yMax));
        root.Controls.Add(BuildCoordGroup("Text2 Window", _t2xMin, _t2yMin, _t2xMax, _t2yMax));
        root.Controls.Add(BuildOutputGroup());
        root.Controls.Add(BuildButtonsGroup());

        _grid.Columns.Add("file_name", "File");
        _grid.Columns.Add("block_name", "Block");
        _grid.Columns.Add("text1", "Text 1");
        _grid.Columns.Add("text2", "Text 2");
        _grid.Columns.Add("insert_x", "X");
        _grid.Columns.Add("insert_y", "Y");
        split.Panel2.Controls.Add(_grid);

        LoadSettings();
        FormClosing += (_, _) => SaveSettings();
    }

    private void TrySetWindowIcon()
    {
        try
        {
            string iconPath = Path.Combine(AppContext.BaseDirectory, "Resources", "app.ico");
            if (File.Exists(iconPath))
            {
                Icon = new Icon(iconPath);
            }
        }
        catch
        {
            // Ignore icon loading errors.
        }
    }

    private GroupBox BuildFileGroup()
    {
        var g = MakeGroup("1) Source Excel");
        var row = Flow();
        var browse = new Button { Text = "...", Width = 35 };
        browse.Click += (_, _) =>
        {
            using var ofd = new OpenFileDialog { Filter = "Excel files (*.xlsx;*.xls)|*.xlsx;*.xls" };
            if (ofd.ShowDialog(this) == DialogResult.OK)
            {
                _excelPath.Text = ofd.FileName;
            }
        };
        row.Controls.Add(_excelPath);
        row.Controls.Add(browse);
        g.Controls.Add(row);
        return g;
    }

    private GroupBox BuildBlockGroup()
    {
        var g = MakeGroup("2) Block Filter");
        g.Controls.Add(Labeled("Block name:", _blockName, _blockNameRegex));
        g.Controls.Add(Labeled("Layer:", _blockLayer, _blockLayerRegex));
        g.Controls.Add(_blockInfo);
        return g;
    }

    private GroupBox BuildTextGroup()
    {
        var g = MakeGroup("3) Text Filter");
        g.Controls.Add(Labeled("Content:", _textContent, _textContentRegex));
        g.Controls.Add(Labeled("Layer:", _textLayer, _textLayerRegex));
        g.Controls.Add(Labeled("Color:", _textColor, _textColorRegex));

        var h = Flow();
        h.Controls.Add(new Label { Text = "Height min:" });
        h.Controls.Add(_heightMin);
        h.Controls.Add(new Label { Text = "max:" });
        h.Controls.Add(_heightMax);
        g.Controls.Add(h);
        g.Controls.Add(_textInfo);
        return g;
    }

    private GroupBox BuildCoordGroup(string title, TextBox xMin, TextBox yMin, TextBox xMax, TextBox yMax)
    {
        var g = MakeGroup(title);
        var r = Flow();
        r.Controls.Add(new Label { Text = "xMin" });
        r.Controls.Add(xMin);
        r.Controls.Add(new Label { Text = "yMin" });
        r.Controls.Add(yMin);
        r.Controls.Add(new Label { Text = "xMax" });
        r.Controls.Add(xMax);
        r.Controls.Add(new Label { Text = "yMax" });
        r.Controls.Add(yMax);
        g.Controls.Add(r);
        return g;
    }

    private GroupBox BuildOutputGroup()
    {
        var g = MakeGroup("4) Output");
        var row = Flow();
        var browse = new Button { Text = "...", Width = 35 };
        browse.Click += (_, _) =>
        {
            using var sfd = new SaveFileDialog { Filter = "Excel files (*.xlsx)|*.xlsx", DefaultExt = "xlsx" };
            if (sfd.ShowDialog(this) == DialogResult.OK)
            {
                _outputPath.Text = sfd.FileName;
            }
        };
        row.Controls.Add(_outputPath);
        row.Controls.Add(browse);
        g.Controls.Add(row);

        var nameRow = Flow();
        nameRow.Controls.Add(new Label { Text = "Sheet name:" });
        nameRow.Controls.Add(_entityName);
        g.Controls.Add(nameRow);
        g.Controls.Add(_appendExisting);
        return g;
    }

    private GroupBox BuildButtonsGroup()
    {
        var g = MakeGroup("5) Actions");
        var row = Flow();
        var match = new Button { Text = "Match", Width = 120 };
        var clear = new Button { Text = "Clear", Width = 120 };
        var save = new Button { Text = "Save", Width = 120 };
        match.Click += (_, _) => RunMatch();
        clear.Click += (_, _) => ClearAll();
        save.Click += (_, _) => SaveMatches();
        row.Controls.Add(match);
        row.Controls.Add(clear);
        row.Controls.Add(save);
        g.Controls.Add(row);
        return g;
    }

    private void RunMatch()
    {
        try
        {
            if (string.IsNullOrWhiteSpace(_excelPath.Text))
            {
                MessageBox.Show(this, "Select source Excel file.", "Match", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var (blocks, texts) = ExcelData.ReadInput(_excelPath.Text.Trim());
            var fBlocks = ApplyBlockFilters(blocks);
            var fTexts = ApplyTextFilters(texts);

            _blockInfo.Text = $"Blocks: {fBlocks.Count} / {blocks.Count}";
            _textInfo.Text = $"Texts: {fTexts.Count} / {texts.Count}";

            var w1 = ReadWindow(_t1xMin, _t1yMin, _t1xMax, _t1yMax);
            var w2 = ReadWindow(_t2xMin, _t2yMin, _t2xMax, _t2yMax);

            _matches.Clear();
            _grid.Rows.Clear();

            foreach (BlockRow b in fBlocks)
            {
                var t1 = fTexts
                    .Where(t => t.FileName == b.FileName && InWindow(t.InsertX, t.InsertY, b.InsertX, b.InsertY, w1))
                    .OrderBy(t => t.InsertX)
                    .Select(t => t.TextPlain ?? string.Empty);
                var t2 = fTexts
                    .Where(t => t.FileName == b.FileName && InWindow(t.InsertX, t.InsertY, b.InsertX, b.InsertY, w2))
                    .OrderBy(t => t.InsertX)
                    .Select(t => t.TextPlain ?? string.Empty);

                var row = new MatchRow(
                    b.FileName,
                    b.BlockName,
                    b.InsertX,
                    b.InsertY,
                    string.Concat(t1),
                    string.Concat(t2));

                _matches.Add(row);
            }

            foreach (var row in _matches.Take(500))
            {
                _grid.Rows.Add(row.FileName, row.BlockName, row.Text1, row.Text2, row.InsertX, row.InsertY);
            }

            SaveSettings();
            MessageBox.Show(this, $"Matched blocks: {_matches.Count}\nPreview rows: {Math.Min(500, _matches.Count)}", "Done");
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void SaveMatches()
    {
        try
        {
            if (_matches.Count == 0)
            {
                MessageBox.Show(this, "No matched data. Click Match first.", "Save", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (string.IsNullOrWhiteSpace(_outputPath.Text))
            {
                MessageBox.Show(this, "Select output .xlsx file.", "Save", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string sheetName = string.IsNullOrWhiteSpace(_entityName.Text) ? "Sheet1" : _entityName.Text.Trim();
            ExcelData.WriteOutput(_outputPath.Text.Trim(), sheetName, _matches, _appendExisting.Checked);
            SaveSettings();

            MessageBox.Show(this, $"Saved {_matches.Count} rows\nFile: {_outputPath.Text}\nSheet: {sheetName}", "Saved");
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void ClearAll()
    {
        foreach (Control c in ControlsRecursive(this))
        {
            if (c is TextBox tb)
            {
                tb.Clear();
            }
            else if (c is CheckBox cb)
            {
                cb.Checked = false;
            }
        }

        _t1xMin.Text = "0";
        _t1yMin.Text = "0";
        _t1xMax.Text = "0";
        _t1yMax.Text = "0";
        _t2xMin.Text = "0";
        _t2yMin.Text = "0";
        _t2xMax.Text = "0";
        _t2yMax.Text = "0";
        _entityName.Text = "Sheet1";

        _matches.Clear();
        _grid.Rows.Clear();
        _blockInfo.Text = "Blocks: 0 / 0";
        _textInfo.Text = "Texts: 0 / 0";
        SaveSettings();
    }

    private List<BlockRow> ApplyBlockFilters(List<BlockRow> blocks)
    {
        return blocks.Where(b =>
            MatchField(b.BlockName, _blockName.Text, _blockNameRegex.Checked) &&
            MatchField(b.Layer, _blockLayer.Text, _blockLayerRegex.Checked)).ToList();
    }

    private List<TextRow> ApplyTextFilters(List<TextRow> texts)
    {
        double? hMin = ParseNullableDouble(_heightMin.Text);
        double? hMax = ParseNullableDouble(_heightMax.Text);

        return texts.Where(t =>
            MatchField(t.TextContent, _textContent.Text, _textContentRegex.Checked) &&
            MatchField(t.Layer, _textLayer.Text, _textLayerRegex.Checked) &&
            MatchField(t.Color, _textColor.Text, _textColorRegex.Checked) &&
            (!hMin.HasValue || t.Height >= hMin.Value) &&
            (!hMax.HasValue || t.Height <= hMax.Value)).ToList();
    }

    private static bool MatchField(string source, string pattern, bool regex)
    {
        pattern = pattern?.Trim() ?? string.Empty;
        source ??= string.Empty;
        if (pattern.Length == 0)
        {
            return true;
        }

        if (!regex)
        {
            return source.Contains(pattern, StringComparison.OrdinalIgnoreCase);
        }

        try
        {
            return Regex.IsMatch(source, $"^(?:{pattern})$", RegexOptions.IgnoreCase);
        }
        catch
        {
            return false;
        }
    }

    private static bool InWindow(double tx, double ty, double bx, double by, CoordWindow w)
    {
        double x = tx - bx;
        double y = ty - by;
        return x >= w.XMin && x <= w.XMax && y >= w.YMin && y <= w.YMax;
    }

    private static CoordWindow ReadWindow(TextBox xMin, TextBox yMin, TextBox xMax, TextBox yMax)
    {
        return new CoordWindow(
            ParseDouble(xMin.Text),
            ParseDouble(yMin.Text),
            ParseDouble(xMax.Text),
            ParseDouble(yMax.Text));
    }

    private static double ParseDouble(string? s)
    {
        return double.TryParse(s, out double value) ? value : 0d;
    }

    private static double? ParseNullableDouble(string? s)
    {
        if (string.IsNullOrWhiteSpace(s))
        {
            return null;
        }

        return double.TryParse(s, out double value) ? value : null;
    }

    private GroupBox MakeGroup(string title)
    {
        return new GroupBox
        {
            Text = title,
            AutoSize = true,
            Dock = DockStyle.Top,
            Padding = new Padding(10),
            Margin = new Padding(0, 0, 0, 8)
        };
    }

    private static FlowLayoutPanel Flow()
    {
        return new FlowLayoutPanel
        {
            AutoSize = true,
            Dock = DockStyle.Top,
            FlowDirection = FlowDirection.LeftToRight,
            WrapContents = false
        };
    }

    private static Control Labeled(string label, Control input, Control optional)
    {
        var row = Flow();
        row.Controls.Add(new Label { Text = label, Width = 80 });
        row.Controls.Add(input);
        row.Controls.Add(optional);
        return row;
    }

    private void LoadSettings()
    {
        if (!File.Exists(_settingsPath))
        {
            return;
        }

        try
        {
            var data = JsonSerializer.Deserialize<AppSettings>(File.ReadAllText(_settingsPath));
            if (data == null)
            {
                return;
            }

            _excelPath.Text = data.ExcelPath ?? "";
            _outputPath.Text = data.OutputPath ?? "";
            _entityName.Text = string.IsNullOrWhiteSpace(data.EntityName) ? "Sheet1" : data.EntityName;
            _appendExisting.Checked = data.AppendExisting;

            _blockName.Text = data.BlockName ?? "";
            _blockLayer.Text = data.BlockLayer ?? "";
            _blockNameRegex.Checked = data.BlockNameRegex;
            _blockLayerRegex.Checked = data.BlockLayerRegex;

            _textContent.Text = data.TextContent ?? "";
            _textLayer.Text = data.TextLayer ?? "";
            _textColor.Text = data.TextColor ?? "";
            _textContentRegex.Checked = data.TextContentRegex;
            _textLayerRegex.Checked = data.TextLayerRegex;
            _textColorRegex.Checked = data.TextColorRegex;
            _heightMin.Text = data.HeightMin ?? "";
            _heightMax.Text = data.HeightMax ?? "";

            _t1xMin.Text = data.T1XMin ?? "0";
            _t1yMin.Text = data.T1YMin ?? "0";
            _t1xMax.Text = data.T1XMax ?? "0";
            _t1yMax.Text = data.T1YMax ?? "0";
            _t2xMin.Text = data.T2XMin ?? "0";
            _t2yMin.Text = data.T2YMin ?? "0";
            _t2xMax.Text = data.T2XMax ?? "0";
            _t2yMax.Text = data.T2YMax ?? "0";
        }
        catch
        {
            // Ignore broken settings file.
        }
    }

    private void SaveSettings()
    {
        var data = new AppSettings
        {
            ExcelPath = _excelPath.Text,
            OutputPath = _outputPath.Text,
            EntityName = _entityName.Text,
            AppendExisting = _appendExisting.Checked,
            BlockName = _blockName.Text,
            BlockLayer = _blockLayer.Text,
            BlockNameRegex = _blockNameRegex.Checked,
            BlockLayerRegex = _blockLayerRegex.Checked,
            TextContent = _textContent.Text,
            TextLayer = _textLayer.Text,
            TextColor = _textColor.Text,
            TextContentRegex = _textContentRegex.Checked,
            TextLayerRegex = _textLayerRegex.Checked,
            TextColorRegex = _textColorRegex.Checked,
            HeightMin = _heightMin.Text,
            HeightMax = _heightMax.Text,
            T1XMin = _t1xMin.Text,
            T1YMin = _t1yMin.Text,
            T1XMax = _t1xMax.Text,
            T1YMax = _t1yMax.Text,
            T2XMin = _t2xMin.Text,
            T2YMin = _t2yMin.Text,
            T2XMax = _t2xMax.Text,
            T2YMax = _t2yMax.Text
        };

        File.WriteAllText(_settingsPath, JsonSerializer.Serialize(data, new JsonSerializerOptions { WriteIndented = true }));
    }

    private static IEnumerable<Control> ControlsRecursive(Control root)
    {
        foreach (Control child in root.Controls)
        {
            yield return child;
            foreach (var nested in ControlsRecursive(child))
            {
                yield return nested;
            }
        }
    }
}

internal sealed class AppSettings
{
    public string? ExcelPath { get; set; }
    public string? OutputPath { get; set; }
    public string? EntityName { get; set; }
    public bool AppendExisting { get; set; }
    public string? BlockName { get; set; }
    public string? BlockLayer { get; set; }
    public bool BlockNameRegex { get; set; }
    public bool BlockLayerRegex { get; set; }
    public string? TextContent { get; set; }
    public string? TextLayer { get; set; }
    public string? TextColor { get; set; }
    public bool TextContentRegex { get; set; }
    public bool TextLayerRegex { get; set; }
    public bool TextColorRegex { get; set; }
    public string? HeightMin { get; set; }
    public string? HeightMax { get; set; }
    public string? T1XMin { get; set; }
    public string? T1YMin { get; set; }
    public string? T1XMax { get; set; }
    public string? T1YMax { get; set; }
    public string? T2XMin { get; set; }
    public string? T2YMin { get; set; }
    public string? T2XMax { get; set; }
    public string? T2YMax { get; set; }
}

internal readonly record struct CoordWindow(double XMin, double YMin, double XMax, double YMax);

internal readonly record struct BlockRow(
    string FileName,
    string Layer,
    string BlockName,
    double InsertX,
    double InsertY);

internal readonly record struct TextRow(
    string FileName,
    string Layer,
    double Height,
    string Color,
    double InsertX,
    double InsertY,
    double Rotation,
    string TextContent,
    string TextPlain);

internal readonly record struct MatchRow(
    string FileName,
    string BlockName,
    double InsertX,
    double InsertY,
    string Text1,
    string Text2);

internal static class ExcelData
{
    public static (List<BlockRow> Blocks, List<TextRow> Texts) ReadInput(string path)
    {
        string ext = Path.GetExtension(path).ToLowerInvariant();
        if (ext == ".xls")
        {
            return ReadInputWithExcelDataReader(path);
        }

        using var wb = new XLWorkbook(path);
        IXLWorksheet blocksSheet = FindSheet(wb, "blocks");
        IXLWorksheet textsSheet = FindSheet(wb, "texts");

        var blockHeaders = ReadHeaderMap(blocksSheet);
        var textHeaders = ReadHeaderMap(textsSheet);

        string[] requiredBlocks = ["file_name", "layer", "block_name", "insert_x", "insert_y"];
        string[] requiredTexts = ["file_name", "layer", "height", "color", "insert_x", "insert_y", "text_content", "text_plain"];
        EnsureRequired(blockHeaders, requiredBlocks, "blocks");
        EnsureRequired(textHeaders, requiredTexts, "texts");

        var blocks = new List<BlockRow>();
        foreach (var row in blocksSheet.RowsUsed().Skip(1))
        {
            string file = GetStr(row, blockHeaders, "file_name");
            string layer = GetStr(row, blockHeaders, "layer");
            string name = GetStr(row, blockHeaders, "block_name");
            double x = GetNum(row, blockHeaders, "insert_x");
            double y = GetNum(row, blockHeaders, "insert_y");
            blocks.Add(new BlockRow(file, layer, name, x, y));
        }

        var texts = new List<TextRow>();
        foreach (var row in textsSheet.RowsUsed().Skip(1))
        {
            string file = GetStr(row, textHeaders, "file_name");
            string layer = GetStr(row, textHeaders, "layer");
            double height = GetNum(row, textHeaders, "height");
            string color = GetStr(row, textHeaders, "color");
            double x = GetNum(row, textHeaders, "insert_x");
            double y = GetNum(row, textHeaders, "insert_y");
            double rotation = textHeaders.ContainsKey("rotation") ? GetNum(row, textHeaders, "rotation") : 0d;
            string content = GetStr(row, textHeaders, "text_content");
            string plain = GetStr(row, textHeaders, "text_plain");
            texts.Add(new TextRow(file, layer, height, color, x, y, rotation, content, plain));
        }

        return (blocks, texts);
    }

    private static (List<BlockRow> Blocks, List<TextRow> Texts) ReadInputWithExcelDataReader(string path)
    {
        using var stream = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        using var reader = ExcelReaderFactory.CreateReader(stream);

        var sheets = new Dictionary<string, List<object?[]>>(StringComparer.OrdinalIgnoreCase);
        do
        {
            var rows = new List<object?[]>();
            while (reader.Read())
            {
                var row = new object?[reader.FieldCount];
                for (int i = 0; i < reader.FieldCount; i++)
                {
                    row[i] = reader.GetValue(i);
                }

                rows.Add(row);
            }

            sheets[reader.Name.Trim()] = rows;
        } while (reader.NextResult());

        if (!TryGetSheet(sheets, "blocks", out var blockRowsRaw))
        {
            throw new InvalidOperationException("Sheet 'blocks' not found.");
        }

        if (!TryGetSheet(sheets, "texts", out var textRowsRaw))
        {
            throw new InvalidOperationException("Sheet 'texts' not found.");
        }

        var blockHeaders = ReadHeaderMap(blockRowsRaw);
        var textHeaders = ReadHeaderMap(textRowsRaw);

        string[] requiredBlocks = ["file_name", "layer", "block_name", "insert_x", "insert_y"];
        string[] requiredTexts = ["file_name", "layer", "height", "color", "insert_x", "insert_y", "text_content", "text_plain"];
        EnsureRequired(blockHeaders, requiredBlocks, "blocks");
        EnsureRequired(textHeaders, requiredTexts, "texts");

        var blocks = new List<BlockRow>();
        foreach (var row in blockRowsRaw.Skip(1))
        {
            blocks.Add(new BlockRow(
                GetStr(row, blockHeaders, "file_name"),
                GetStr(row, blockHeaders, "layer"),
                GetStr(row, blockHeaders, "block_name"),
                GetNum(row, blockHeaders, "insert_x"),
                GetNum(row, blockHeaders, "insert_y")));
        }

        var texts = new List<TextRow>();
        foreach (var row in textRowsRaw.Skip(1))
        {
            texts.Add(new TextRow(
                GetStr(row, textHeaders, "file_name"),
                GetStr(row, textHeaders, "layer"),
                GetNum(row, textHeaders, "height"),
                GetStr(row, textHeaders, "color"),
                GetNum(row, textHeaders, "insert_x"),
                GetNum(row, textHeaders, "insert_y"),
                textHeaders.ContainsKey("rotation") ? GetNum(row, textHeaders, "rotation") : 0d,
                GetStr(row, textHeaders, "text_content"),
                GetStr(row, textHeaders, "text_plain")));
        }

        return (blocks, texts);
    }

    public static void WriteOutput(string path, string sheetName, IReadOnlyList<MatchRow> rows, bool appendExisting)
    {
        XLWorkbook wb;
        if (appendExisting && File.Exists(path))
        {
            wb = new XLWorkbook(path);
        }
        else
        {
            wb = new XLWorkbook();
        }

        using (wb)
        {
            if (wb.Worksheets.TryGetWorksheet(sheetName, out var oldSheet))
            {
                oldSheet.Delete();
            }

            var ws = wb.Worksheets.Add(sheetName);
            ws.Cell(1, 1).Value = "filename";
            ws.Cell(1, 2).Value = "minx";
            ws.Cell(1, 3).Value = "miny";
            ws.Cell(1, 4).Value = "text1";
            ws.Cell(1, 5).Value = "text2";

            int r = 2;
            foreach (MatchRow row in rows)
            {
                ws.Cell(r, 1).Value = row.FileName;
                ws.Cell(r, 2).Value = row.InsertX;
                ws.Cell(r, 3).Value = row.InsertY;
                ws.Cell(r, 4).Value = row.Text1;
                ws.Cell(r, 5).Value = row.Text2;
                r++;
            }

            wb.SaveAs(path);
        }
    }

    private static IXLWorksheet FindSheet(XLWorkbook wb, string name)
    {
        foreach (IXLWorksheet sheet in wb.Worksheets)
        {
            if (sheet.Name.Equals(name, StringComparison.OrdinalIgnoreCase))
            {
                return sheet;
            }
        }

        throw new InvalidOperationException($"Sheet '{name}' not found.");
    }

    private static Dictionary<string, int> ReadHeaderMap(IXLWorksheet ws)
    {
        var map = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        foreach (IXLCell cell in ws.Row(1).CellsUsed())
        {
            string key = cell.GetString().Trim().ToLowerInvariant();
            if (!string.IsNullOrWhiteSpace(key))
            {
                map[key] = cell.Address.ColumnNumber;
            }
        }

        return map;
    }

    private static Dictionary<string, int> ReadHeaderMap(List<object?[]> rows)
    {
        var map = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        if (rows.Count == 0)
        {
            return map;
        }

        object?[] header = rows[0];
        for (int i = 0; i < header.Length; i++)
        {
            string key = (header[i]?.ToString() ?? string.Empty).Trim().ToLowerInvariant();
            if (!string.IsNullOrWhiteSpace(key))
            {
                map[key] = i;
            }
        }

        return map;
    }

    private static void EnsureRequired(Dictionary<string, int> map, IEnumerable<string> cols, string sheet)
    {
        foreach (string col in cols)
        {
            if (!map.ContainsKey(col))
            {
                throw new InvalidOperationException($"Column '{col}' is missing on sheet '{sheet}'.");
            }
        }
    }

    private static string GetStr(IXLRow row, IReadOnlyDictionary<string, int> map, string key)
    {
        return row.Cell(map[key]).GetString();
    }

    private static double GetNum(IXLRow row, IReadOnlyDictionary<string, int> map, string key)
    {
        var cell = row.Cell(map[key]);
        return cell.TryGetValue<double>(out double value) ? value : 0d;
    }

    private static bool TryGetSheet(Dictionary<string, List<object?[]>> sheets, string wanted, out List<object?[]> rows)
    {
        foreach (var kv in sheets)
        {
            if (kv.Key.Equals(wanted, StringComparison.OrdinalIgnoreCase))
            {
                rows = kv.Value;
                return true;
            }
        }

        rows = [];
        return false;
    }

    private static string GetStr(object?[] row, IReadOnlyDictionary<string, int> map, string key)
    {
        if (!map.TryGetValue(key, out int index) || index >= row.Length)
        {
            return string.Empty;
        }

        return row[index]?.ToString() ?? string.Empty;
    }

    private static double GetNum(object?[] row, IReadOnlyDictionary<string, int> map, string key)
    {
        if (!map.TryGetValue(key, out int index) || index >= row.Length)
        {
            return 0d;
        }

        object? val = row[index];
        if (val == null)
        {
            return 0d;
        }

        if (val is double d)
        {
            return d;
        }

        if (double.TryParse(val.ToString(), out double parsed))
        {
            return parsed;
        }

        return 0d;
    }
}
