using System;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace XCelReportEngine.Tests
{
    internal sealed class WorkbookFixture : IDisposable
    {
        internal WorkbookFixture(SpreadsheetDocumentType type = SpreadsheetDocumentType.Workbook)
        {
            DirectoryPath = Path.Combine(Path.GetTempPath(), "XCelReportEngine.Tests", Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(DirectoryPath);
            SourcePath = Path.Combine(DirectoryPath, type == SpreadsheetDocumentType.Template ? "source.xltx" : "source.xlsx");
            OutputPath = Path.Combine(DirectoryPath, "output.xlsx");
            ImagePath = Path.Combine(DirectoryPath, "pixel.png");
            File.WriteAllBytes(ImagePath, Convert.FromBase64String(
                "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Y9ZVdAAAAAASUVORK5CYII="));
            CreateWorkbook(SourcePath, type);
        }

        internal string DirectoryPath { get; }

        internal string SourcePath { get; }

        internal string OutputPath { get; }

        internal string ImagePath { get; }

        public void Dispose()
        {
            try
            {
                Directory.Delete(DirectoryPath, true);
            }
            catch
            {
                // Test cleanup only.
            }
        }

        private static void CreateWorkbook(string path, SpreadsheetDocumentType type)
        {
            using (var document = SpreadsheetDocument.Create(path, type))
            {
                var workbookPart = document.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();
                var sharedStrings = workbookPart.AddNewPart<SharedStringTablePart>();
                sharedStrings.SharedStringTable = new SharedStringTable(new SharedStringItem(new Text("shared text")));
                var styles = workbookPart.AddNewPart<WorkbookStylesPart>();
                styles.Stylesheet = new Stylesheet(
                    new Fonts(new Font()) { Count = 1 },
                    new Fills(
                        new Fill(new PatternFill { PatternType = PatternValues.None }),
                        new Fill(new PatternFill { PatternType = PatternValues.Gray125 })) { Count = 2 },
                    new Borders(new Border()) { Count = 1 },
                    new CellStyleFormats(new CellFormat()) { Count = 1 },
                    new CellFormats(new CellFormat()) { Count = 1 },
                    new CellStyles(new CellStyle { Name = "Normal", FormatId = 0, BuiltinId = 0 }) { Count = 1 });
                var sheets = workbookPart.Workbook.AppendChild(new Sheets());
                AddWorksheet(workbookPart, sheets, "First", 1, true);
                AddWorksheet(workbookPart, sheets, "Second", 2);
                workbookPart.Workbook.Save();
            }
        }

        private static void AddWorksheet(WorkbookPart workbookPart, Sheets sheets, string name, uint sheetId, bool addTestCells = false)
        {
            var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            var sheetData = new SheetData();
            if (addTestCells)
            {
                sheetData.Append(new Row(
                    new Cell { CellReference = "A1", DataType = CellValues.SharedString, CellValue = new CellValue("0"), StyleIndex = 0 },
                    new Cell { CellReference = "B1", CellFormula = new CellFormula("1+1"), CellValue = new CellValue("2"), StyleIndex = 0 },
                    new Cell { CellReference = "C1", DataType = CellValues.Boolean, CellValue = new CellValue("1") })
                { RowIndex = 1 });
            }

            worksheetPart.Worksheet = new Worksheet(sheetData);
            worksheetPart.Worksheet.Save();
            sheets.Append(new Sheet
            {
                Id = workbookPart.GetIdOfPart(worksheetPart),
                SheetId = new UInt32Value(sheetId),
                Name = new StringValue(name)
            });
        }
    }
}
