using System;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace IzAutoCADPlugin
{
    public sealed class ExportDialog : Form
    {
        private readonly ListBox _filesList;
        private readonly TextBox _outputFileTextBox;
        private readonly ProgressBar _filesProgressBar;
        private readonly ProgressBar _entitiesProgressBar;
        private readonly Label _filesProgressLabel;
        private readonly Label _entitiesProgressLabel;
        private readonly Label _statusLabel;
        private readonly Button _runButton;
        private readonly Button _addFilesButton;
        private readonly Button _clearFilesButton;
        private readonly Button _browseOutputButton;

        public ExportDialog()
        {
            Text = "DWG Export to Excel";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterParent;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            ClientSize = new Size(760, 520);

            _addFilesButton = new Button
            {
                Text = "Add DWG Files",
                Location = new Point(20, 20),
                Size = new Size(140, 32)
            };
            _addFilesButton.Click += (_, __) => AddFiles();

            _clearFilesButton = new Button
            {
                Text = "Clear List",
                Location = new Point(170, 20),
                Size = new Size(140, 32)
            };
            _clearFilesButton.Click += (_, __) => _filesList.Items.Clear();

            _filesList = new ListBox
            {
                Location = new Point(20, 64),
                Size = new Size(720, 220),
                HorizontalScrollbar = true
            };

            var outputLabel = new Label
            {
                Text = "Output File:",
                Location = new Point(20, 302),
                Size = new Size(140, 20)
            };

            _outputFileTextBox = new TextBox
            {
                Location = new Point(20, 326),
                Size = new Size(600, 26)
            };

            _browseOutputButton = new Button
            {
                Text = "Browse...",
                Location = new Point(630, 324),
                Size = new Size(110, 30)
            };
            _browseOutputButton.Click += (_, __) => BrowseOutputFile();

            _runButton = new Button
            {
                Text = "Run",
                Location = new Point(20, 370),
                Size = new Size(140, 36)
            };
            _runButton.Click += (_, __) => RunExport();

            _filesProgressLabel = new Label
            {
                Text = "Files: 0/0",
                Location = new Point(20, 420),
                Size = new Size(200, 20)
            };

            _filesProgressBar = new ProgressBar
            {
                Location = new Point(20, 444),
                Size = new Size(720, 20),
                Minimum = 0,
                Maximum = 1,
                Value = 0
            };

            _entitiesProgressLabel = new Label
            {
                Text = "Entities: blocks 0, texts 0",
                Location = new Point(20, 470),
                Size = new Size(360, 20)
            };

            _entitiesProgressBar = new ProgressBar
            {
                Location = new Point(20, 494),
                Size = new Size(720, 20),
                Style = ProgressBarStyle.Marquee,
                MarqueeAnimationSpeed = 20
            };

            _statusLabel = new Label
            {
                Text = "Ready.",
                Location = new Point(180, 378),
                Size = new Size(560, 20)
            };

            Controls.Add(_addFilesButton);
            Controls.Add(_clearFilesButton);
            Controls.Add(_filesList);
            Controls.Add(outputLabel);
            Controls.Add(_outputFileTextBox);
            Controls.Add(_browseOutputButton);
            Controls.Add(_runButton);
            Controls.Add(_statusLabel);
            Controls.Add(_filesProgressLabel);
            Controls.Add(_filesProgressBar);
            Controls.Add(_entitiesProgressLabel);
            Controls.Add(_entitiesProgressBar);
        }

        private void AddFiles()
        {
            using (var dialog = new OpenFileDialog())
            {
                dialog.Filter = "AutoCAD Drawing (*.dwg)|*.dwg";
                dialog.Multiselect = true;
                dialog.Title = "Select DWG Files";
                if (dialog.ShowDialog(this) != DialogResult.OK)
                {
                    return;
                }

                foreach (string path in dialog.FileNames)
                {
                    if (_filesList.Items.Contains(path))
                    {
                        continue;
                    }

                    _filesList.Items.Add(path);
                }

                if (string.IsNullOrWhiteSpace(_outputFileTextBox.Text) && dialog.FileNames.Length > 0)
                {
                    string baseDir = Path.GetDirectoryName(dialog.FileNames[0]) ?? string.Empty;
                    _outputFileTextBox.Text = Path.Combine(baseDir, "result.xlsx");
                }
            }
        }

        private void BrowseOutputFile()
        {
            using (var dialog = new SaveFileDialog())
            {
                dialog.Filter = "Excel Workbook (*.xlsx)|*.xlsx";
                dialog.DefaultExt = "xlsx";
                dialog.AddExtension = true;
                dialog.Title = "Select output Excel file";
                dialog.FileName = string.IsNullOrWhiteSpace(_outputFileTextBox.Text)
                    ? "result.xlsx"
                    : Path.GetFileName(_outputFileTextBox.Text);
                string initialDir = Path.GetDirectoryName(_outputFileTextBox.Text);
                if (!string.IsNullOrWhiteSpace(initialDir) && Directory.Exists(initialDir))
                {
                    dialog.InitialDirectory = initialDir;
                }

                if (dialog.ShowDialog(this) == DialogResult.OK)
                {
                    _outputFileTextBox.Text = dialog.FileName;
                }
            }
        }

        private void RunExport()
        {
            var files = _filesList.Items.Cast<string>().ToList();
            if (files.Count == 0)
            {
                MessageBox.Show(this, "Add at least one DWG file.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string excelFile = _outputFileTextBox.Text?.Trim();
            if (string.IsNullOrWhiteSpace(excelFile))
            {
                MessageBox.Show(this, "Select an output Excel file.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (!excelFile.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
            {
                excelFile += ".xlsx";
                _outputFileTextBox.Text = excelFile;
            }

            string outputDir = Path.GetDirectoryName(excelFile);
            if (string.IsNullOrWhiteSpace(outputDir))
            {
                MessageBox.Show(this, "Output path is invalid.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            ToggleUi(false);

            var progress = new ExportProgress
            {
                TotalFiles = files.Count
            };
            UpdateProgressUi(progress);
            _statusLabel.Text = "Export is running...";
            Application.DoEvents();

            try
            {
                var result = DwgCsvExporter.Export(files, excelFile, progressUpdate =>
                {
                    UpdateProgressUi(progressUpdate);
                    Application.DoEvents();
                });

                string message = $"Done.\n\nBlocks: {result.BlocksWritten}\nTexts: {result.TextsWritten}\nExcel: {result.ExcelPath}";
                if (result.FailedFiles.Count > 0)
                {
                    message += $"\n\nErrors ({result.FailedFiles.Count}):\n" + string.Join("\n", result.FailedFiles);
                }

                _statusLabel.Text = "Export completed.";
                MessageBox.Show(this, message, "Result", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                _statusLabel.Text = "Export failed.";
                MessageBox.Show(this, ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                ToggleUi(true);
            }
        }

        private void ToggleUi(bool enabled)
        {
            _runButton.Enabled = enabled;
            _addFilesButton.Enabled = enabled;
            _clearFilesButton.Enabled = enabled;
            _browseOutputButton.Enabled = enabled;
            _filesList.Enabled = enabled;
            _outputFileTextBox.Enabled = enabled;
        }

        private void UpdateProgressUi(ExportProgress progress)
        {
            _filesProgressBar.Maximum = Math.Max(progress.TotalFiles, 1);
            _filesProgressBar.Value = Math.Max(0, Math.Min(progress.ProcessedFiles, _filesProgressBar.Maximum));
            _filesProgressLabel.Text = $"Files: {progress.ProcessedFiles}/{progress.TotalFiles}";
            _entitiesProgressLabel.Text = $"Entities: blocks {progress.ProcessedBlocks}, texts {progress.ProcessedTexts}";
        }
    }
}
