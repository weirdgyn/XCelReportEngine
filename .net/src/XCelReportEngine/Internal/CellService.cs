using System;
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace XCelReportEngine.Internal
{
    internal static class CellService
    {
        internal static string ReadAsString(WorkbookSession session, string address, string operation)
        {
            var normalizedAddress = CellAddress.Normalize(address, operation);
            var cell = FindCell(session, normalizedAddress);
            if (cell == null)
            {
                return string.Empty;
            }

            if (cell.CellFormula != null)
            {
                throw new ReportEngineException(
                    ReportErrorCode.FormulaReadNotSupported,
                    operation,
                    $"Cell {normalizedAddress} contains a formula. Formula results must be calculated by the LabVIEW application.");
            }

            return ReadCellValue(session, cell, normalizedAddress, operation);
        }

        internal static void WriteString(WorkbookSession session, string address, string value, string operation)
        {
            var cell = GetOrCreateCell(session, CellAddress.Normalize(address, operation));
            ClearValue(cell);
            cell.DataType = CellValues.InlineString;
            cell.InlineString = new InlineString(new Text(value ?? string.Empty) { Space = SpaceProcessingModeValues.Preserve });
            GetWorksheet(session).Save();
        }

        internal static void WriteDouble(WorkbookSession session, string address, double value, string operation)
        {
            if (double.IsNaN(value) || double.IsInfinity(value))
            {
                throw new ReportEngineException(ReportErrorCode.InvalidArgument, operation, "Excel numeric cells do not support NaN or infinity.");
            }

            var cell = GetOrCreateCell(session, CellAddress.Normalize(address, operation));
            ClearValue(cell);
            cell.DataType = CellValues.Number;
            cell.CellValue = new CellValue(value.ToString("R", CultureInfo.InvariantCulture));
            GetWorksheet(session).Save();
        }

        internal static void WriteBoolean(WorkbookSession session, string address, bool value, string operation)
        {
            var cell = GetOrCreateCell(session, CellAddress.Normalize(address, operation));
            ClearValue(cell);
            cell.DataType = CellValues.Boolean;
            cell.CellValue = new CellValue(value ? "1" : "0");
            GetWorksheet(session).Save();
        }

        internal static string[] ReadStringRange(
            WorkbookSession session,
            int startRowIndex,
            int startColumnIndex,
            int rowCount,
            int columnCount,
            string operation)
        {
            var elementCount = CellAddress.ValidateRange(startRowIndex, startColumnIndex, rowCount, columnCount, operation);
            var result = new string[elementCount];
            var formulas = new System.Collections.Generic.List<string>();
            var cellsByAddress = GetWorksheet(session)
                .GetFirstChild<SheetData>()?
                .Descendants<Cell>()
                .Where(cell => !string.IsNullOrEmpty(cell.CellReference?.Value))
                .GroupBy(cell => cell.CellReference!.Value!, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(group => group.Key, group => group.First(), StringComparer.OrdinalIgnoreCase)
                ?? new System.Collections.Generic.Dictionary<string, Cell>(StringComparer.OrdinalIgnoreCase);

            for (var rowOffset = 0; rowOffset < rowCount; rowOffset++)
            {
                for (var columnOffset = 0; columnOffset < columnCount; columnOffset++)
                {
                    var address = CellAddress.FromZeroBased(startRowIndex + rowOffset, startColumnIndex + columnOffset, operation);
                    cellsByAddress.TryGetValue(address, out var cell);
                    if (cell?.CellFormula != null)
                    {
                        formulas.Add(address);
                        continue;
                    }

                    result[rowOffset * columnCount + columnOffset] = cell == null
                        ? string.Empty
                        : ReadCellValue(session, cell, address, operation);
                }
            }

            if (formulas.Count > 0)
            {
                var shown = string.Join(", ", formulas.Take(10));
                var suffix = formulas.Count > 10 ? $" (+{formulas.Count - 10} more)" : string.Empty;
                throw new ReportEngineException(
                    ReportErrorCode.FormulaReadNotSupported,
                    operation,
                    $"The requested range contains formula cells: {shown}{suffix}. Formula results must be calculated by the LabVIEW application.");
            }

            return result;
        }

        internal static void WriteStringRange(
            WorkbookSession session,
            int startRowIndex,
            int startColumnIndex,
            int rowCount,
            int columnCount,
            string[] values,
            string operation)
        {
            var expectedLength = CellAddress.ValidateRange(startRowIndex, startColumnIndex, rowCount, columnCount, operation);
            if (values == null)
            {
                throw new ReportEngineException(ReportErrorCode.InvalidArgument, operation, "The values array is required.");
            }

            if (values.Length != expectedLength)
            {
                throw new ReportEngineException(
                    ReportErrorCode.InvalidRangeDimensions,
                    operation,
                    $"The values array contains {values.Length} elements; {expectedLength} are required for a {rowCount} x {columnCount} range.");
            }

            for (var rowOffset = 0; rowOffset < rowCount; rowOffset++)
            {
                for (var columnOffset = 0; columnOffset < columnCount; columnOffset++)
                {
                    var address = CellAddress.FromZeroBased(startRowIndex + rowOffset, startColumnIndex + columnOffset, operation);
                    var cell = GetOrCreateCell(session, address);
                    ClearValue(cell);
                    cell.DataType = CellValues.InlineString;
                    var value = values[rowOffset * columnCount + columnOffset] ?? string.Empty;
                    cell.InlineString = new InlineString(new Text(value) { Space = SpaceProcessingModeValues.Preserve });
                }
            }

            GetWorksheet(session).Save();
        }

        private static Cell? FindCell(WorkbookSession session, string address)
        {
            return GetWorksheet(session)
                .GetFirstChild<SheetData>()?
                .Descendants<Cell>()
                .FirstOrDefault(item => string.Equals(item.CellReference?.Value, address, StringComparison.OrdinalIgnoreCase));
        }

        private static string ReadCellValue(WorkbookSession session, Cell cell, string address, string operation)
        {
            if (cell.DataType == null)
            {
                return cell.CellValue?.Text ?? string.Empty;
            }

            var dataType = cell.DataType.Value;
            if (dataType == CellValues.SharedString)
            {
                return ReadSharedString(session, cell, address, operation);
            }

            if (dataType == CellValues.InlineString)
            {
                return cell.InlineString?.InnerText ?? string.Empty;
            }

            if (dataType == CellValues.Boolean)
            {
                return string.Equals(cell.CellValue?.Text, "1", StringComparison.Ordinal) ? "TRUE" : "FALSE";
            }

            return cell.CellValue?.Text ?? cell.InnerText ?? string.Empty;
        }

        internal static Cell GetOrCreateCell(WorkbookSession session, string address)
        {
            var worksheet = GetWorksheet(session);
            var sheetData = worksheet.GetFirstChild<SheetData>();
            sheetData ??= worksheet.PrependChild(new SheetData());

            CellAddress.Parse(address, out var rowNumber, out var columnNumber);
            var row = sheetData.Elements<Row>().FirstOrDefault(item => item.RowIndex?.Value == rowNumber);
            if (row == null)
            {
                row = new Row { RowIndex = rowNumber };
                var followingRow = sheetData.Elements<Row>().FirstOrDefault(item => item.RowIndex?.Value > rowNumber);
                if (followingRow == null)
                {
                    sheetData.Append(row);
                }
                else
                {
                    sheetData.InsertBefore(row, followingRow);
                }
            }

            var existing = row.Elements<Cell>().FirstOrDefault(item => string.Equals(item.CellReference?.Value, address, StringComparison.OrdinalIgnoreCase));
            if (existing != null)
            {
                return existing;
            }

            var cell = new Cell { CellReference = address };
            var followingCell = row.Elements<Cell>().FirstOrDefault(item => CellAddress.GetColumnNumber(item.CellReference?.Value) > columnNumber);
            if (followingCell == null)
            {
                row.Append(cell);
            }
            else
            {
                row.InsertBefore(cell, followingCell);
            }

            return cell;
        }

        private static void ClearValue(Cell cell)
        {
            cell.CellFormula?.Remove();
            cell.CellValue?.Remove();
            cell.InlineString?.Remove();
        }

        private static string ReadSharedString(WorkbookSession session, Cell cell, string address, string operation)
        {
            if (!int.TryParse(cell.CellValue?.Text, NumberStyles.None, CultureInfo.InvariantCulture, out var index))
            {
                throw new ReportEngineException(ReportErrorCode.InvalidWorkbook, operation, $"Cell {address} has an invalid shared-string index.");
            }

            var items = session.WorkbookPart.SharedStringTablePart?.SharedStringTable?.Elements<SharedStringItem>().ToArray();
            if (items == null || index < 0 || index >= items.Length)
            {
                throw new ReportEngineException(ReportErrorCode.InvalidWorkbook, operation, $"Cell {address} references a missing shared string.");
            }

            return items[index].InnerText;
        }

        private static Worksheet GetWorksheet(WorkbookSession session)
        {
            return session.ActiveWorksheetPart.Worksheet
                ?? throw new ReportEngineException(ReportErrorCode.InvalidWorkbook, "ActiveWorksheet", "The worksheet root element is missing.");
        }

    }
}
