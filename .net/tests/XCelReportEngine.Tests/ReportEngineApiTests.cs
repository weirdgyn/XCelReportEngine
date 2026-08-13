using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using Xunit;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace XCelReportEngine.Tests
{
    public sealed class ReportEngineApiTests
    {
        [Theory]
        [InlineData(0, 0, "A1")]
        [InlineData(19, 17, "R20")]
        [InlineData(1048575, 16383, "XFD1048576")]
        public void CellCoordinateConversion_RoundTripsZeroBasedIndexes(int rowIndex, int columnIndex, string expectedAddress)
        {
            Assert.Equal(expectedAddress, ReportEngineApi.ConvertCellIndexToAddress(rowIndex, columnIndex));
            Assert.Equal([rowIndex, columnIndex], ReportEngineApi.ConvertCellAddressToIndex(expectedAddress));
        }

        [Fact]
        public void CellCoordinateConversion_NormalizesAbsoluteLowercaseAddress()
        {
            Assert.Equal([12, 14], ReportEngineApi.ConvertCellAddressToIndex("$o$13"));
        }

        [Fact]
        public void OpenWorkbook_ExposesWorksheetNamesAndSelection()
        {
            using var fixture = new WorkbookFixture();
            using var api = ReportEngineApi.Create();
            var sessionId = api.OpenWorkbook(fixture.SourcePath, fixture.OutputPath);

            Assert.Equal(["First", "Second"], api.GetWorksheetNames(sessionId));
            Assert.Equal("First", api.GetActiveWorksheetName(sessionId));

            api.SelectWorksheetByName(sessionId, "Second");
            Assert.Equal("Second", api.GetActiveWorksheetName(sessionId));

            api.SelectWorksheetByIndex(sessionId, 0);
            Assert.Equal("First", api.GetActiveWorksheetName(sessionId));
            api.CloseWorkbook(sessionId, false);
        }

        [Fact]
        public void SaveWorkbook_PublishesValidXlsxWithoutClosingSession()
        {
            using var fixture = new WorkbookFixture();
            using var api = ReportEngineApi.Create();
            var sessionId = api.OpenWorkbook(fixture.SourcePath, fixture.OutputPath);
            api.SelectWorksheetByName(sessionId, "Second");
            api.SaveWorkbook(sessionId);

            Assert.True(File.Exists(fixture.OutputPath));
            using (var saved = SpreadsheetDocument.Open(fixture.OutputPath, false))
            {
                Assert.Equal(SpreadsheetDocumentType.Workbook, saved.DocumentType);
                Assert.Equal(2, saved.WorkbookPart!.Workbook!.Sheets?.Elements<Sheet>().Count());
                AssertValid(saved);
            }

            Assert.Equal("Second", api.GetActiveWorksheetName(sessionId));
            api.CloseWorkbook(sessionId, false);
        }

        [Fact]
        public void OpenTemplate_MaterializesWorkbookContentType()
        {
            using var fixture = new WorkbookFixture(SpreadsheetDocumentType.Template);
            using var api = ReportEngineApi.Create();
            var sessionId = api.OpenWorkbook(fixture.SourcePath, fixture.OutputPath);
            api.CloseWorkbook(sessionId, true);

            using var saved = SpreadsheetDocument.Open(fixture.OutputPath, false);
            Assert.Equal(SpreadsheetDocumentType.Workbook, saved.DocumentType);
            AssertValid(saved);
        }

        [Fact]
        public void LockAndUnlock_AreIdempotent()
        {
            using var fixture = new WorkbookFixture();
            using var api = ReportEngineApi.Create();
            var sessionId = api.OpenWorkbook(fixture.SourcePath, fixture.OutputPath);
            api.LockAllWorksheets(sessionId, "test-key");
            api.LockAllWorksheets(sessionId, "test-key");
            api.SaveWorkbook(sessionId);

            using (var locked = SpreadsheetDocument.Open(fixture.OutputPath, false))
            {
                Assert.All(
                    locked.WorkbookPart!.WorksheetParts,
                    part => Assert.Single(part.Worksheet!.Elements<SheetProtection>()));
                AssertValid(locked);
            }

            api.UnlockAllWorksheets(sessionId);
            api.CloseWorkbook(sessionId, true);

            using var unlocked = SpreadsheetDocument.Open(fixture.OutputPath, false);
            Assert.All(
                unlocked.WorkbookPart!.WorksheetParts,
                part => Assert.Empty(part.Worksheet!.Elements<SheetProtection>()));
            AssertValid(unlocked);
        }

        [Fact]
        public void UnknownSession_ProducesStableErrorCode()
        {
            using var api = ReportEngineApi.Create();
            var exception = Assert.Throws<ReportEngineException>(() => api.GetWorksheetNames(999));
            Assert.Equal((int)ReportErrorCode.SessionNotFound, exception.ErrorCode);
            Assert.Equal("SessionNotFound", exception.ErrorCodeName);
        }

        [Fact]
        public void ValidateSession_AcceptsAnOpenSession()
        {
            using var fixture = new WorkbookFixture();
            using var api = ReportEngineApi.Create();
            var sessionId = api.OpenWorkbook(fixture.SourcePath, fixture.OutputPath);

            api.ValidateSession(sessionId);
        }

        [Fact]
        public void ValidateSession_RejectsUnknownAndClosedSessions()
        {
            using var fixture = new WorkbookFixture();
            using var api = ReportEngineApi.Create();
            var unknown = Assert.Throws<ReportEngineException>(() => api.ValidateSession(999));
            Assert.Equal((int)ReportErrorCode.SessionNotFound, unknown.ErrorCode);
            Assert.Equal("Session", unknown.Operation);

            var sessionId = api.OpenWorkbook(fixture.SourcePath, fixture.OutputPath);
            api.CloseWorkbook(sessionId, false);

            var closed = Assert.Throws<ReportEngineException>(() => api.ValidateSession(sessionId));
            Assert.Equal((int)ReportErrorCode.SessionNotFound, closed.ErrorCode);
            Assert.Equal("Session", closed.Operation);
        }

        [Fact]
        public void CellRead_HandlesSharedStringsBooleansBlanksAndRejectsFormulas()
        {
            using var fixture = new WorkbookFixture();
            using var api = ReportEngineApi.Create();
            var sessionId = api.OpenWorkbook(fixture.SourcePath, fixture.OutputPath);

            Assert.Equal("shared text", api.ReadCellStringByAddress(sessionId, "$a$1"));
            Assert.Equal("TRUE", api.ReadCellStringByIndex(sessionId, 0, 2));
            Assert.Equal(string.Empty, api.ReadCellStringByAddress(sessionId, "D20"));

            var exception = Assert.Throws<ReportEngineException>(() => api.ReadCellStringByAddress(sessionId, "B1"));
            Assert.Equal((int)ReportErrorCode.FormulaReadNotSupported, exception.ErrorCode);
            Assert.Equal("ReadCellStringByAddress", exception.Operation);
        }

        [Fact]
        public void CellWrite_RoundTripsTypedValuesAndPreservesExistingStyle()
        {
            using var fixture = new WorkbookFixture();
            using var api = ReportEngineApi.Create();
            var sessionId = api.OpenWorkbook(fixture.SourcePath, fixture.OutputPath);

            api.WriteCellStringByAddress(sessionId, "A1", "  updated  ");
            api.WriteCellDoubleByIndex(sessionId, 1, 1, 1234.5);
            api.WriteCellBooleanByAddress(sessionId, "C2", false);

            Assert.Equal("  updated  ", api.ReadCellStringByIndex(sessionId, 0, 0));
            Assert.Equal("1234.5", api.ReadCellStringByAddress(sessionId, "B2"));
            Assert.Equal("FALSE", api.ReadCellStringByIndex(sessionId, 1, 2));

            api.CloseWorkbook(sessionId, true);
            using var saved = SpreadsheetDocument.Open(fixture.OutputPath, false);
            var firstPart = saved.WorkbookPart!.WorksheetParts.First();
            var cell = firstPart.Worksheet!.Descendants<Cell>().Single(item => item.CellReference?.Value == "A1");
            Assert.Equal(0U, cell.StyleIndex?.Value);
            AssertValid(saved);
        }

        [Theory]
        [InlineData("A0")]
        [InlineData("XFE1")]
        [InlineData("A1:B2")]
        public void CellAddress_InvalidAddressProducesStableError(string address)
        {
            using var fixture = new WorkbookFixture();
            using var api = ReportEngineApi.Create();
            var sessionId = api.OpenWorkbook(fixture.SourcePath, fixture.OutputPath);
            var exception = Assert.Throws<ReportEngineException>(() => api.ReadCellStringByAddress(sessionId, address));
            Assert.Equal((int)ReportErrorCode.InvalidCellAddress, exception.ErrorCode);
        }

        [Fact]
        public void StringRange_RoundTripsFlatRowMajorArray()
        {
            using var fixture = new WorkbookFixture();
            using var api = ReportEngineApi.Create();
            var sessionId = api.OpenWorkbook(fixture.SourcePath, fixture.OutputPath);
            var values = new[] { "r0c0", "r0c1", "r0c2", "r1c0", "", "r1c2" };

            api.WriteStringRangeByIndex(sessionId, 1, 3, 2, 3, values);

            Assert.Equal(values, api.ReadStringRangeByIndex(sessionId, 1, 3, 2, 3));
            Assert.Equal("r0c0", api.ReadCellStringByAddress(sessionId, "D2"));
            Assert.Equal("r1c2", api.ReadCellStringByAddress(sessionId, "F3"));
            api.CloseWorkbook(sessionId, true);

            using var saved = SpreadsheetDocument.Open(fixture.OutputPath, false);
            AssertValid(saved);
        }

        [Fact]
        public void StringRange_FormulaAndDimensionErrorsAreDeterministic()
        {
            using var fixture = new WorkbookFixture();
            using var api = ReportEngineApi.Create();
            var sessionId = api.OpenWorkbook(fixture.SourcePath, fixture.OutputPath);

            var formulaError = Assert.Throws<ReportEngineException>(() => api.ReadStringRangeByIndex(sessionId, 0, 0, 1, 3));
            Assert.Equal((int)ReportErrorCode.FormulaReadNotSupported, formulaError.ErrorCode);
            Assert.Contains("B1", formulaError.Message);

            var sizeError = Assert.Throws<ReportEngineException>(() =>
                api.WriteStringRangeByIndex(sessionId, 2, 2, 2, 2, ["only one"]));
            Assert.Equal((int)ReportErrorCode.InvalidRangeDimensions, sizeError.ErrorCode);
        }

        [Fact]
        public void AppendImage_UsesCellAddressPrecedenceAndDeduplicatesMedia()
        {
            using WorkbookFixture fixture = new();
            using var api = ReportEngineApi.Create();
            var sessionId = api.OpenWorkbook(fixture.SourcePath, fixture.OutputPath);

            Assert.Equal(0, api.AppendImage(sessionId, fixture.ImagePath, 4, 2, string.Empty, 0, string.Empty));
            Assert.Equal(1, api.AppendImage(sessionId, fixture.ImagePath, 0, 0, "H7", 1, "ignored caption"));
            api.CloseWorkbook(sessionId, true);

            using var saved = SpreadsheetDocument.Open(fixture.OutputPath, false);
            var drawingPart = saved.WorkbookPart!.WorksheetParts.First().DrawingsPart!;
            Assert.Single(drawingPart.ImageParts);
            var anchors = drawingPart.WorksheetDrawing!.Elements<Xdr.TwoCellAnchor>().ToArray();
            Assert.Equal(2, anchors.Length);
            Assert.Equal("2", anchors[0].FromMarker!.ColumnId!.Text);
            Assert.Equal("4", anchors[0].FromMarker!.RowId!.Text);
            Assert.Equal("7", anchors[1].FromMarker!.ColumnId!.Text);
            Assert.Equal("6", anchors[1].FromMarker!.RowId!.Text);
            Assert.All(anchors, anchor => Assert.Equal(Xdr.EditAsValues.OneCell, anchor.EditAs?.Value));

            using (var embedded = drawingPart.ImageParts.Single().GetStream())
            using (var expected = File.OpenRead(fixture.ImagePath))
            {
                Assert.Equal(expected.Length, embedded.Length);
            }

            AssertValid(saved);
        }

        [Fact]
        public void AppendImage_InvalidInputProducesStableErrors()
        {
            using var fixture = new WorkbookFixture();
            using var api = ReportEngineApi.Create();
            var sessionId = api.OpenWorkbook(fixture.SourcePath, fixture.OutputPath);
            var missing = Assert.Throws<ReportEngineException>(() =>
                api.AppendImage(sessionId, Path.Combine(fixture.DirectoryPath, "missing.png"), 0, 0, string.Empty, 0, string.Empty));
            Assert.Equal((int)ReportErrorCode.ImageFileNotFound, missing.ErrorCode);

            var address = Assert.Throws<ReportEngineException>(() =>
                api.AppendImage(sessionId, fixture.ImagePath, 0, 0, "BAD", 0, string.Empty));
            Assert.Equal((int)ReportErrorCode.InvalidCellAddress, address.ErrorCode);
        }

        [Fact]
        public void FormatImage_ScalesLastPictureAndAppliesGrayscale()
        {
            using var fixture = new WorkbookFixture();
            using var api = ReportEngineApi.Create();
            var sessionId = api.OpenWorkbook(fixture.SourcePath, fixture.OutputPath);
            api.AppendImage(sessionId, fixture.ImagePath, 0, 0, string.Empty, 0, string.Empty);
            api.FormatImage(sessionId, 1, 0.5, -1, -1, -1, 2);
            api.CloseWorkbook(sessionId, true);

            using var saved = SpreadsheetDocument.Open(fixture.OutputPath, false);
            var picture = saved.WorkbookPart!.WorksheetParts.First().DrawingsPart!
                .WorksheetDrawing!.Descendants<Xdr.Picture>().Single();
            Assert.Equal(4763L, picture.ShapeProperties!.Transform2D!.Extents!.Cx!.Value);
            Assert.Equal(4763L, picture.ShapeProperties.Transform2D.Extents.Cy!.Value);
            Assert.Single(picture.BlipFill!.Blip!.Elements<DocumentFormat.OpenXml.Drawing.Grayscale>());
            AssertValid(saved);
        }

        [Fact]
        public void FormatImage_UsesQuantizedMetricDimensionsAndValidatesIndex()
        {
            using var fixture = new WorkbookFixture();
            using var api = ReportEngineApi.Create();
            var sessionId = api.OpenWorkbook(fixture.SourcePath, fixture.OutputPath);
            api.AppendImage(sessionId, fixture.ImagePath, 0, 0, string.Empty, 0, string.Empty);
            api.FormatImage(sessionId, 1, 1, 0, 2.54, 5.08, 0);
            api.SaveWorkbook(sessionId);

            using (var saved = SpreadsheetDocument.Open(fixture.OutputPath, false))
            {
                var extents = saved.WorkbookPart!.WorksheetParts.First().DrawingsPart!
                    .WorksheetDrawing!.Descendants<Xdr.Picture>().Single().ShapeProperties!.Transform2D!.Extents!;
                Assert.Equal(914400L, extents.Cy!.Value);
                Assert.Equal(1828800L, extents.Cx!.Value);
                AssertValid(saved);
            }

            var exception = Assert.Throws<ReportEngineException>(() => api.FormatImage(sessionId, 0, 1, 4, -1, -1, 0));
            Assert.Equal((int)ReportErrorCode.PictureIndexOutOfRange, exception.ErrorCode);
        }

        [Fact]
        public void SetCellAlignment_AppliesRangeAndPreservesCellValues()
        {
            using var fixture = new WorkbookFixture();
            using var api = ReportEngineApi.Create();
            var sessionId = api.OpenWorkbook(fixture.SourcePath, fixture.OutputPath);
            api.SetCellAlignmentByAddress(sessionId, "A1", "B2", 2, 2, true, 45);
            api.CloseWorkbook(sessionId, true);

            using var saved = SpreadsheetDocument.Open(fixture.OutputPath, false);
            var worksheet = saved.WorkbookPart!.WorksheetParts.First().Worksheet!;
            var cells = worksheet.Descendants<Cell>()
                .Where(cell => new[] { "A1", "B1", "A2", "B2" }.Contains(cell.CellReference?.Value))
                .ToArray();
            Assert.Equal(4, cells.Length);
            Assert.Equal("0", cells.Single(cell => cell.CellReference?.Value == "A1").CellValue?.Text);

            var formats = saved.WorkbookPart.WorkbookStylesPart!.Stylesheet!.CellFormats!;
            Assert.All(cells, cell =>
            {
                var format = formats.Elements<CellFormat>().ElementAt((int)cell.StyleIndex!.Value);
                Assert.Equal(HorizontalAlignmentValues.Center, format.Alignment!.Horizontal!.Value);
                Assert.Equal(VerticalAlignmentValues.Top, format.Alignment.Vertical!.Value);
                Assert.True(format.Alignment.WrapText!.Value);
                Assert.Equal(45U, format.Alignment.TextRotation!.Value);
            });
            AssertValid(saved);
        }

        [Fact]
        public void SetCellColorAndBorder_CanApplyFillAndBorderIndependently()
        {
            using var fixture = new WorkbookFixture();
            using var api = ReportEngineApi.Create();
            var sessionId = api.OpenWorkbook(fixture.SourcePath, fixture.OutputPath);
            api.SetCellColorAndBorderByIndex(sessionId, 0, 0, 1, 2, true, 0x12AB34, false, 0, 0, 0);
            api.SetCellColorAndBorderByAddress(sessionId, "A1", "B1", false, 0, true, 0xCC3300, 1, 15);
            api.CloseWorkbook(sessionId, true);

            using var saved = SpreadsheetDocument.Open(fixture.OutputPath, false);
            var workbookStyles = saved.WorkbookPart!.WorkbookStylesPart!.Stylesheet!;
            var cells = saved.WorkbookPart.WorksheetParts.First().Worksheet!.Descendants<Cell>()
                .Where(cell => cell.CellReference?.Value == "A1" || cell.CellReference?.Value == "B1")
                .ToArray();
            Assert.Equal(2, cells.Length);
            Assert.All(cells, cell =>
            {
                var format = workbookStyles.CellFormats!.Elements<CellFormat>().ElementAt((int)cell.StyleIndex!.Value);
                var fill = workbookStyles.Fills!.Elements<Fill>().ElementAt((int)format.FillId!.Value);
                Assert.Equal("FF12AB34", fill.PatternFill!.ForegroundColor!.Rgb!.Value);
                var border = workbookStyles.Borders!.Elements<Border>().ElementAt((int)format.BorderId!.Value);
                Assert.Equal(BorderStyleValues.Thin, border.LeftBorder!.Style!.Value);
                Assert.Equal("FFCC3300", border.BottomBorder!.Color!.Rgb!.Value);
            });
            AssertValid(saved);
        }

        [Fact]
        public void CellFormatting_InvalidArgumentsProduceStableErrors()
        {
            using var fixture = new WorkbookFixture();
            using var api = ReportEngineApi.Create();
            var sessionId = api.OpenWorkbook(fixture.SourcePath, fixture.OutputPath);
            Assert.Equal((int)ReportErrorCode.InvalidArgument,
                Assert.Throws<ReportEngineException>(() => api.SetCellAlignmentByAddress(sessionId, "A1", "A1", 99, 0, false, 0)).ErrorCode);
            Assert.Equal((int)ReportErrorCode.InvalidRangeDimensions,
                Assert.Throws<ReportEngineException>(() => api.SetCellColorAndBorderByAddress(sessionId, "B2", "A1", true, 0, false, 0, 0, 0)).ErrorCode);
            Assert.Equal((int)ReportErrorCode.InvalidArgument,
                Assert.Throws<ReportEngineException>(() => api.SetCellColorAndBorderByAddress(sessionId, "A1", "A1", true, -1, false, 0, 0, 0)).ErrorCode);
        }

        private static void AssertValid(SpreadsheetDocument document)
        {
            var errors = new OpenXmlValidator().Validate(document).ToArray();
            Assert.True(errors.Length == 0, string.Join("\n", errors.Select(error => error.Description)));
        }
    }
}
