using System.Text.Json.Nodes;

namespace IzArchicadPlacer.App;

internal sealed class MainForm : Form
{
    private readonly TextBox _excelPath = new() { Width = 420 };
    private readonly TextBox _tapirUrl = new() { Width = 220 };
    private readonly ComboBox _units = new() { Width = 80, DropDownStyle = ComboBoxStyle.DropDownList };

    private readonly TextBox _prop1Set = new() { Width = 180 };
    private readonly TextBox _prop1Name = new() { Width = 140 };
    private readonly TextBox _prop2Set = new() { Width = 180 };
    private readonly TextBox _prop2Name = new() { Width = 140 };
    private readonly CheckBox _writeToId = new() { Text = "Write text1 to Element ID" };

    private readonly DataGridView _gridFiles = new()
    {
        Dock = DockStyle.Fill,
        AllowUserToAddRows = false,
        AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
    };

    private readonly DataGridView _gridSheets = new()
    {
        Dock = DockStyle.Fill,
        AllowUserToAddRows = false,
        AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
    };

    private readonly List<SheetData> _excelSheets = [];
    private readonly List<StoryInfo> _stories = [];
    private readonly Dictionary<string, List<double>> _filenameToLevels = new(StringComparer.OrdinalIgnoreCase);
    private AppSettings _settings;

    public MainForm()
    {
        Text = "IZ Archicad Placer (AC26 + Tapir)";
        Width = 1200;
        Height = 800;
        StartPosition = FormStartPosition.CenterScreen;
        TrySetWindowIcon();
        _settings = SettingsStore.Load();

        var root = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            RowCount = 4,
            ColumnCount = 1
        };
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 55));
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 45));
        Controls.Add(root);

        root.Controls.Add(BuildTopPanel(), 0, 0);
        root.Controls.Add(BuildPropsPanel(), 0, 1);
        root.Controls.Add(BuildFileTablePanel(), 0, 2);
        root.Controls.Add(BuildSheetTablePanel(), 0, 3);

        _gridFiles.CellDoubleClick += GridFilesOnCellDoubleClick;

        LoadSettingsToUi();
    }

    private void TrySetWindowIcon()
    {
        try
        {
            string exePath = Application.ExecutablePath;
            if (File.Exists(exePath))
            {
                Icon = Icon.ExtractAssociatedIcon(exePath);
            }
        }
        catch
        {
            // Ignore icon loading errors.
        }
    }

    private Control BuildTopPanel()
    {
        var panel = new FlowLayoutPanel
        {
            AutoSize = true,
            Dock = DockStyle.Top,
            Padding = new Padding(8)
        };

        var pickExcel = new Button { Text = "Select Excel" };
        pickExcel.Click += (_, _) => SelectExcel();

        var reloadStories = new Button { Text = "Reload Stories" };
        reloadStories.Click += async (_, _) => await ReloadStoriesAsync();

        var place = new Button { Text = "Place Items" };
        place.Click += async (_, _) => await PlaceItemsAsync();

        var browseExcel = new Button { Text = "..." };
        browseExcel.Click += (_, _) => SelectExcel();

        _units.Items.AddRange(["mm", "cm", "m"]);

        panel.Controls.Add(new Label { Text = "Excel:" });
        panel.Controls.Add(_excelPath);
        panel.Controls.Add(browseExcel);
        panel.Controls.Add(new Label { Text = "Tapir URL:" });
        panel.Controls.Add(_tapirUrl);
        panel.Controls.Add(new Label { Text = "Units:" });
        panel.Controls.Add(_units);
        panel.Controls.Add(pickExcel);
        panel.Controls.Add(reloadStories);
        panel.Controls.Add(place);

        return panel;
    }

    private Control BuildPropsPanel()
    {
        var group = new GroupBox
        {
            Text = "Property Mapping",
            Dock = DockStyle.Top,
            AutoSize = true,
            Padding = new Padding(8)
        };

        var t = new TableLayoutPanel
        {
            Dock = DockStyle.Top,
            AutoSize = true,
            ColumnCount = 6
        };
        group.Controls.Add(t);

        t.Controls.Add(new Label { Text = "text1 -> set" }, 0, 0);
        t.Controls.Add(_prop1Set, 1, 0);
        t.Controls.Add(new Label { Text = "name" }, 2, 0);
        t.Controls.Add(_prop1Name, 3, 0);
        t.Controls.Add(_writeToId, 4, 0);
        t.SetColumnSpan(_writeToId, 2);

        t.Controls.Add(new Label { Text = "text2 -> set" }, 0, 1);
        t.Controls.Add(_prop2Set, 1, 1);
        t.Controls.Add(new Label { Text = "name" }, 2, 1);
        t.Controls.Add(_prop2Name, 3, 1);

        return group;
    }

    private Control BuildFileTablePanel()
    {
        var group = new GroupBox
        {
            Text = "Filenames -> Floors",
            Dock = DockStyle.Fill
        };
        _gridFiles.Columns.Add("filename", "Filename");
        _gridFiles.Columns.Add("floors", "Floors (e.g. 1,2,-1)");
        group.Controls.Add(_gridFiles);
        return group;
    }

    private Control BuildSheetTablePanel()
    {
        var group = new GroupBox
        {
            Text = "Sheets -> Library part",
            Dock = DockStyle.Fill
        };
        _gridSheets.Columns.Add("sheet", "Sheet");
        _gridSheets.Columns.Add("library", "Library part name");
        group.Controls.Add(_gridSheets);
        return group;
    }

    private void LoadSettingsToUi()
    {
        _tapirUrl.Text = string.IsNullOrWhiteSpace(_settings.TapirBaseUrl) ? "http://127.0.0.1:19725" : _settings.TapirBaseUrl;
        _units.SelectedItem = _settings.Units;
        if (_units.SelectedIndex < 0)
        {
            _units.SelectedIndex = 0;
        }

        _prop1Set.Text = _settings.Prop1Set;
        _prop1Name.Text = _settings.Prop1Name;
        _prop2Set.Text = _settings.Prop2Set;
        _prop2Name.Text = _settings.Prop2Name;
        _writeToId.Checked = _settings.WriteToId;
    }

    private void SaveSettingsFromUi()
    {
        _settings.Units = _units.SelectedItem?.ToString() ?? "mm";
        _settings.Prop1Set = _prop1Set.Text;
        _settings.Prop1Name = _prop1Name.Text;
        _settings.Prop2Set = _prop2Set.Text;
        _settings.Prop2Name = _prop2Name.Text;
        _settings.WriteToId = _writeToId.Checked;
        _settings.TapirBaseUrl = _tapirUrl.Text.Trim();
        SettingsStore.Save(_settings);
    }

    private void SelectExcel()
    {
        using var ofd = new OpenFileDialog { Filter = "Excel files (*.xlsx;*.xls)|*.xlsx;*.xls" };
        if (ofd.ShowDialog(this) != DialogResult.OK)
        {
            return;
        }

        _excelPath.Text = ofd.FileName;
        try
        {
            _excelSheets.Clear();
            _excelSheets.AddRange(ExcelInputReader.ReadAll(ofd.FileName));

            var files = new SortedSet<string>(StringComparer.OrdinalIgnoreCase);
            _gridFiles.Rows.Clear();
            _gridSheets.Rows.Clear();

            foreach (var sheet in _excelSheets)
            {
                _gridSheets.Rows.Add(sheet.SheetName, string.Empty);
                foreach (var row in sheet.Rows)
                {
                    files.Add(row.Filename);
                }
            }

            foreach (string f in files)
            {
                _gridFiles.Rows.Add(f, string.Empty);
            }

            MessageBox.Show(this, $"Loaded sheets: {_excelSheets.Count}\nFiles: {files.Count}", "Excel");
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, ex.Message, "Excel Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private async Task ReloadStoriesAsync()
    {
        try
        {
            SaveSettingsFromUi();
            var client = new TapirClient(_tapirUrl.Text.Trim());
            _stories.Clear();
            _stories.AddRange(await client.GetStoriesAsync());
            MessageBox.Show(this, $"Stories loaded: {_stories.Count}", "Stories");
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, ex.Message, "Tapir Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void GridFilesOnCellDoubleClick(object? sender, DataGridViewCellEventArgs e)
    {
        if (e.RowIndex < 0 || e.ColumnIndex != 1)
        {
            return;
        }

        if (_stories.Count == 0)
        {
            MessageBox.Show(this, "Load stories first (Reload Stories).", "Stories");
            return;
        }

        var current = _gridFiles.Rows[e.RowIndex].Cells[1].Value?.ToString() ?? string.Empty;
        using var dlg = new FloorSelectForm(_stories, current);
        if (dlg.ShowDialog(this) == DialogResult.OK)
        {
            _gridFiles.Rows[e.RowIndex].Cells[1].Value = dlg.Result;
        }
    }

    private async Task PlaceItemsAsync()
    {
        try
        {
            SaveSettingsFromUi();
            if (_excelSheets.Count == 0)
            {
                MessageBox.Show(this, "Select Excel first.", "Place");
                return;
            }

            if (_stories.Count == 0)
            {
                MessageBox.Show(this, "Load stories first.", "Place");
                return;
            }

            BuildFilenameToLevels();
            if (_filenameToLevels.Count == 0)
            {
                MessageBox.Show(this, "No floors selected for files.", "Place");
                return;
            }

            var sheetToLibrary = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (DataGridViewRow row in _gridSheets.Rows)
            {
                if (row.IsNewRow) continue;
                string sheet = row.Cells[0].Value?.ToString() ?? string.Empty;
                string lib = row.Cells[1].Value?.ToString() ?? string.Empty;
                sheetToLibrary[sheet] = lib.Trim();
            }

            double scale = _units.SelectedItem?.ToString() switch
            {
                "mm" => 1.0 / 1000.0,
                "cm" => 1.0 / 100.0,
                _ => 1.0
            };

            var payload = new List<ObjectPayload>();
            foreach (var sheet in _excelSheets)
            {
                string lib = sheetToLibrary.TryGetValue(sheet.SheetName, out string? val) ? val : string.Empty;
                if (string.IsNullOrWhiteSpace(lib))
                {
                    continue;
                }

                foreach (var r in sheet.Rows)
                {
                    if (!_filenameToLevels.TryGetValue(r.Filename, out List<double>? levels))
                    {
                        continue;
                    }

                    foreach (double z in levels)
                    {
                        payload.Add(new ObjectPayload
                        {
                            LibraryPartName = lib,
                            X = r.MinX * scale,
                            Y = r.MinY * scale,
                            Z = z,
                            Text1 = string.IsNullOrWhiteSpace(r.Text1) ? null : r.Text1,
                            Text2 = string.IsNullOrWhiteSpace(r.Text2) ? null : r.Text2
                        });
                    }
                }
            }

            if (payload.Count == 0)
            {
                MessageBox.Show(this, "No valid rows to place.", "Place");
                return;
            }

            var client = new TapirClient(_tapirUrl.Text.Trim());
            List<string> guids = await client.CreateObjectsAsync(payload);
            if (guids.Count == 0)
            {
                MessageBox.Show(this, "Tapir did not return element GUIDs.", "Place", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            int written = await WritePropertiesAsync(client, guids, payload);
            MessageBox.Show(this, $"Created elements: {guids.Count}\nProperty writes: {written}", "Done");
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, ex.Message, "Place Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private async Task<int> WritePropertiesAsync(TapirClient client, List<string> guids, List<ObjectPayload> payload)
    {
        List<JsonObject> allProps = await client.GetAllPropertiesAsync();
        string? guid1 = FindPropGuid(allProps, _prop1Set.Text, _prop1Name.Text);
        string? guid2 = FindPropGuid(allProps, _prop2Set.Text, _prop2Name.Text);
        string? idPropGuid = null;
        if (_writeToId.Checked)
        {
            idPropGuid = await client.GetBuiltInElementIdPropertyGuidAsync();
        }

        var writes = new JsonArray();
        for (int i = 0; i < guids.Count && i < payload.Count; i++)
        {
            string elementGuid = guids[i];
            ObjectPayload src = payload[i];

            if (_writeToId.Checked)
            {
                if (!string.IsNullOrWhiteSpace(src.Text1) && !string.IsNullOrWhiteSpace(idPropGuid))
                {
                    writes.Add(CreateWriteItem(elementGuid, idPropGuid, src.Text1));
                }
            }
            else if (!string.IsNullOrWhiteSpace(src.Text1) && !string.IsNullOrWhiteSpace(guid1))
            {
                writes.Add(CreateWriteItem(elementGuid, guid1, src.Text1));
            }

            if (!string.IsNullOrWhiteSpace(src.Text2) && !string.IsNullOrWhiteSpace(guid2))
            {
                writes.Add(CreateWriteItem(elementGuid, guid2, src.Text2));
            }
        }

        if (writes.Count > 0)
        {
            await client.SetPropertyValuesOfElementsAsync(writes);
        }

        return writes.Count;
    }

    private static JsonObject CreateWriteItem(string elementGuid, string propGuid, string value)
    {
        return new JsonObject
        {
            ["elementId"] = new JsonObject { ["guid"] = elementGuid },
            ["propertyId"] = new JsonObject { ["guid"] = propGuid },
            ["propertyValue"] = new JsonObject { ["value"] = value }
        };
    }

    private static string? FindPropGuid(IEnumerable<JsonObject> allProps, string groupName, string propName)
    {
        foreach (JsonObject p in allProps)
        {
            string g = p["propertyGroupName"]?.GetValue<string>() ?? string.Empty;
            string n = p["propertyName"]?.GetValue<string>() ?? string.Empty;
            if (n.Equals(propName, StringComparison.OrdinalIgnoreCase) &&
                (string.IsNullOrWhiteSpace(groupName) || g.Equals(groupName, StringComparison.OrdinalIgnoreCase)))
            {
                return p["propertyId"]?["guid"]?.GetValue<string>();
            }
        }

        return null;
    }

    private void BuildFilenameToLevels()
    {
        _filenameToLevels.Clear();
        var storyByIndex = _stories.ToDictionary(s => s.Index);

        foreach (DataGridViewRow row in _gridFiles.Rows)
        {
            if (row.IsNewRow) continue;
            string filename = row.Cells[0].Value?.ToString() ?? string.Empty;
            string floors = row.Cells[1].Value?.ToString() ?? string.Empty;
            if (string.IsNullOrWhiteSpace(filename) || string.IsNullOrWhiteSpace(floors))
            {
                continue;
            }

            var levels = new List<double>();
            var seen = new HashSet<int>();
            foreach (string token in floors.Split(',', ' ', ';', '|'))
            {
                if (!int.TryParse(token.Trim(), out int userIndex) || userIndex == 0)
                {
                    continue;
                }

                int archIndex = userIndex > 0 ? userIndex - 1 : userIndex;
                if (seen.Contains(archIndex))
                {
                    continue;
                }

                if (storyByIndex.TryGetValue(archIndex, out StoryInfo? story))
                {
                    levels.Add(story.Level);
                    seen.Add(archIndex);
                }
            }

            if (levels.Count > 0)
            {
                _filenameToLevels[filename] = levels;
            }
        }
    }
}

internal sealed class FloorSelectForm : Form
{
    private readonly ListBox _list = new() { Dock = DockStyle.Fill, SelectionMode = SelectionMode.MultiExtended };
    private readonly List<StoryInfo> _stories;
    public string Result { get; private set; } = string.Empty;

    public FloorSelectForm(List<StoryInfo> stories, string current)
    {
        _stories = stories;
        Text = "Select Floors";
        Width = 360;
        Height = 420;
        StartPosition = FormStartPosition.CenterParent;

        var root = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 2 };
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        Controls.Add(root);

        root.Controls.Add(_list, 0, 0);
        var buttons = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true };
        var ok = new Button { Text = "OK", Width = 120 };
        var cancel = new Button { Text = "Cancel", Width = 120 };
        ok.Click += (_, _) => Accept();
        cancel.Click += (_, _) => DialogResult = DialogResult.Cancel;
        buttons.Controls.Add(ok);
        buttons.Controls.Add(cancel);
        root.Controls.Add(buttons, 0, 1);

        foreach (StoryInfo s in _stories)
        {
            _list.Items.Add($"{s.DisplayIndex} - {s.Name}");
        }

        var selected = current.Split(',', ' ', ';').Select(x => x.Trim()).Where(x => x.Length > 0).ToHashSet();
        for (int i = 0; i < _stories.Count; i++)
        {
            if (selected.Contains(_stories[i].DisplayIndex.ToString()))
            {
                _list.SetSelected(i, true);
            }
        }
    }

    private void Accept()
    {
        var chosen = new List<int>();
        foreach (int i in _list.SelectedIndices)
        {
            chosen.Add(_stories[i].DisplayIndex);
        }

        Result = string.Join(", ", chosen.Distinct());
        DialogResult = DialogResult.OK;
    }
}
