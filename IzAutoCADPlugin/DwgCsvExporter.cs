using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using AcadColor = Autodesk.AutoCAD.Colors.Color;
using SheetCell = DocumentFormat.OpenXml.Spreadsheet.Cell;
using SheetRow = DocumentFormat.OpenXml.Spreadsheet.Row;

namespace IzAutoCADPlugin
{
    public sealed class ExportProgress
    {
        public int TotalFiles { get; set; }
        public int ProcessedFiles { get; set; }
        public int ProcessedBlocks { get; set; }
        public int ProcessedTexts { get; set; }
    }

    public sealed class ExportResult
    {
        public int BlocksWritten { get; set; }
        public int TextsWritten { get; set; }
        public string ExcelPath { get; set; }
        public List<string> FailedFiles { get; } = new List<string>();
    }

    public static class DwgCsvExporter
    {
        public static ExportResult Export(
            IReadOnlyList<string> dwgFiles,
            string excelPath,
            Action<ExportProgress> onProgress)
        {
            var progress = new ExportProgress
            {
                TotalFiles = dwgFiles.Count
            };
            var result = new ExportResult
            {
                ExcelPath = excelPath
            };
            var blockRows = new List<BlockRow>();
            var textRows = new List<TextRow>();

            Directory.CreateDirectory(Path.GetDirectoryName(excelPath) ?? ".");

            foreach (string filePath in dwgFiles)
            {
                try
                {
                    ProcessSingleFile(filePath, blockRows, textRows, progress, result, onProgress);
                }
                catch (Exception ex)
                {
                    result.FailedFiles.Add($"{Path.GetFileName(filePath)}: {ex.Message}");
                }
                finally
                {
                    progress.ProcessedFiles++;
                    onProgress?.Invoke(progress);
                }
            }

            WriteExcel(excelPath, blockRows, textRows);
            return result;
        }

        private static void ProcessSingleFile(
            string filePath,
            List<BlockRow> blockRows,
            List<TextRow> textRows,
            ExportProgress progress,
            ExportResult result,
            Action<ExportProgress> onProgress)
        {
            string fileName = Path.GetFileName(filePath);

            using (var db = new Database(false, true))
            {
                db.ReadDwgFile(filePath, FileOpenMode.OpenForReadAndAllShare, true, null);
                db.CloseInput(true);

                using (var tr = db.TransactionManager.StartOpenCloseTransaction())
                {
                    var blockTable = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                    var modelSpaceId = blockTable[BlockTableRecord.ModelSpace];
                    var modelSpace = (BlockTableRecord)tr.GetObject(modelSpaceId, OpenMode.ForRead);
                    foreach (ObjectId entityId in modelSpace)
                    {
                        var entity = tr.GetObject(entityId, OpenMode.ForRead) as Entity;
                        if (entity == null)
                        {
                            continue;
                        }

                        if (entity is BlockReference blockReference)
                        {
                            string blockName = GetBlockName(blockReference, tr);
                            blockRows.Add(new BlockRow
                            {
                                FileName = fileName,
                                Layer = blockReference.Layer,
                                BlockName = blockName,
                                InsertX = FormatDouble(blockReference.Position.X),
                                InsertY = FormatDouble(blockReference.Position.Y)
                            });

                            progress.ProcessedBlocks++;
                            result.BlocksWritten++;
                            onProgress?.Invoke(progress);
                            continue;
                        }

                        if (entity is DBText dbText)
                        {
                            textRows.Add(new TextRow
                            {
                                FileName = fileName,
                                Layer = dbText.Layer,
                                Height = FormatDouble(dbText.Height),
                                Color = GetColorValue(dbText.Color),
                                InsertX = FormatDouble(dbText.Position.X),
                                InsertY = FormatDouble(dbText.Position.Y),
                                Rotation = FormatDouble(ToDegrees(dbText.Rotation)),
                                TextContent = dbText.TextString,
                                TextPlain = dbText.TextString
                            });

                            progress.ProcessedTexts++;
                            result.TextsWritten++;
                            onProgress?.Invoke(progress);
                            continue;
                        }

                        if (entity is MText mText)
                        {
                            textRows.Add(new TextRow
                            {
                                FileName = fileName,
                                Layer = mText.Layer,
                                Height = FormatDouble(mText.TextHeight),
                                Color = GetColorValue(mText.Color),
                                InsertX = FormatDouble(mText.Location.X),
                                InsertY = FormatDouble(mText.Location.Y),
                                Rotation = FormatDouble(ToDegrees(mText.Rotation)),
                                TextContent = mText.Contents,
                                TextPlain = mText.Text
                            });

                            progress.ProcessedTexts++;
                            result.TextsWritten++;
                            onProgress?.Invoke(progress);
                        }
                    }

                    tr.Commit();
                }
            }
        }

        private static void WriteExcel(string excelPath, List<BlockRow> blockRows, List<TextRow> textRows)
        {
            if (File.Exists(excelPath))
            {
                File.Delete(excelPath);
            }

            using (SpreadsheetDocument document = SpreadsheetDocument.Create(excelPath, SpreadsheetDocumentType.Workbook))
            {
                WorkbookPart workbookPart = document.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();

                WorksheetPart blocksPart = workbookPart.AddNewPart<WorksheetPart>();
                WorksheetPart textsPart = workbookPart.AddNewPart<WorksheetPart>();

                blocksPart.Worksheet = new Worksheet(new SheetData());
                textsPart.Worksheet = new Worksheet(new SheetData());

                var sheets = workbookPart.Workbook.AppendChild(new Sheets());
                sheets.Append(new Sheet { Id = workbookPart.GetIdOfPart(blocksPart), SheetId = 1, Name = "blocks" });
                sheets.Append(new Sheet { Id = workbookPart.GetIdOfPart(textsPart), SheetId = 2, Name = "texts" });

                WriteBlocksSheet(blocksPart.Worksheet.GetFirstChild<SheetData>(), blockRows);
                WriteTextsSheet(textsPart.Worksheet.GetFirstChild<SheetData>(), textRows);

                workbookPart.Workbook.Save();
            }
        }

        private static void WriteBlocksSheet(SheetData sheetData, List<BlockRow> rows)
        {
            sheetData.Append(CreateRow("file_name", "layer", "block_name", "insert_x", "insert_y"));
            foreach (var row in rows)
            {
                sheetData.Append(CreateRow(row.FileName, row.Layer, row.BlockName, row.InsertX, row.InsertY));
            }
        }

        private static void WriteTextsSheet(SheetData sheetData, List<TextRow> rows)
        {
            sheetData.Append(CreateRow("file_name", "layer", "height", "color", "insert_x", "insert_y", "rotation", "text_content", "text_plain"));
            foreach (var row in rows)
            {
                sheetData.Append(CreateRow(
                    row.FileName,
                    row.Layer,
                    row.Height,
                    row.Color,
                    row.InsertX,
                    row.InsertY,
                    row.Rotation,
                    row.TextContent,
                    row.TextPlain));
            }
        }

        private static SheetRow CreateRow(params string[] values)
        {
            var row = new SheetRow();
            foreach (string value in values)
            {
                row.Append(new SheetCell
                {
                    DataType = CellValues.String,
                    CellValue = new CellValue(value ?? string.Empty)
                });
            }

            return row;
        }

        private static string GetBlockName(BlockReference blockReference, Transaction tr)
        {
            ObjectId id = blockReference.IsDynamicBlock
                ? blockReference.DynamicBlockTableRecord
                : blockReference.BlockTableRecord;

            var blockDef = tr.GetObject(id, OpenMode.ForRead) as BlockTableRecord;
            return blockDef?.Name ?? string.Empty;
        }

        private static string GetColorValue(AcadColor color)
        {
            if (color == null)
            {
                return string.Empty;
            }

            if (color.IsByLayer)
            {
                return "ByLayer";
            }

            if (color.IsByBlock)
            {
                return "ByBlock";
            }

            if (color.IsByAci)
            {
                return color.ColorIndex.ToString(CultureInfo.InvariantCulture);
            }

            if (color.IsByColor)
            {
                return $"RGB({color.Red},{color.Green},{color.Blue})";
            }

            return color.ColorNameForDisplay;
        }

        private static double ToDegrees(double radians)
        {
            return radians * 180.0 / Math.PI;
        }

        private static string FormatDouble(double value)
        {
            return value.ToString("0.######", CultureInfo.InvariantCulture);
        }

        private sealed class BlockRow
        {
            public string FileName { get; set; }
            public string Layer { get; set; }
            public string BlockName { get; set; }
            public string InsertX { get; set; }
            public string InsertY { get; set; }
        }

        private sealed class TextRow
        {
            public string FileName { get; set; }
            public string Layer { get; set; }
            public string Height { get; set; }
            public string Color { get; set; }
            public string InsertX { get; set; }
            public string InsertY { get; set; }
            public string Rotation { get; set; }
            public string TextContent { get; set; }
            public string TextPlain { get; set; }
        }
    }
}
