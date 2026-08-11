using System;
using System.Globalization;
using System.Linq;
using System.Text;

namespace XCelReportEngine.Internal
{
    internal static class CellAddress
    {
        internal const int MaximumRowIndex = 1048575;
        internal const int MaximumColumnIndex = 16383;

        internal static string FromZeroBased(int rowIndex, int columnIndex, string operation)
        {
            ValidateCoordinates(rowIndex, columnIndex, operation);
            var columnNumber = columnIndex + 1;
            var columnName = new StringBuilder();
            while (columnNumber > 0)
            {
                columnNumber--;
                columnName.Insert(0, (char)('A' + columnNumber % 26));
                columnNumber /= 26;
            }

            return columnName + (rowIndex + 1).ToString(CultureInfo.InvariantCulture);
        }

        internal static string FromOneBased(uint rowNumber, uint columnNumber, string operation)
        {
            if (rowNumber == 0 || columnNumber == 0 || rowNumber > MaximumRowIndex + 1U || columnNumber > MaximumColumnIndex + 1U)
            {
                throw new ReportEngineException(ReportErrorCode.InvalidCellAddress, operation, "Cell coordinates are outside the Excel worksheet limits.");
            }

            return FromZeroBased(checked((int)rowNumber - 1), checked((int)columnNumber - 1), operation);
        }

        internal static string Normalize(string address, string operation)
        {
            if (string.IsNullOrWhiteSpace(address))
            {
                throw Invalid(address, operation);
            }

            var candidate = address.Trim().Replace("$", string.Empty).ToUpperInvariant();
            var separator = candidate.TakeWhile(character => character >= 'A' && character <= 'Z').Count();
            if (separator == 0 || separator == candidate.Length || separator > 3)
            {
                throw Invalid(address, operation);
            }

            var columnText = candidate.Substring(0, separator);
            var rowText = candidate.Substring(separator);
            if (!uint.TryParse(rowText, NumberStyles.None, CultureInfo.InvariantCulture, out var rowNumber)
                || rowNumber == 0
                || rowNumber > MaximumRowIndex + 1)
            {
                throw Invalid(address, operation);
            }

            var columnNumber = GetColumnNumber(columnText);
            if (columnNumber == 0 || columnNumber > MaximumColumnIndex + 1)
            {
                throw Invalid(address, operation);
            }

            return columnText + rowNumber.ToString(CultureInfo.InvariantCulture);
        }

        internal static void ToCoordinates(string address, string operation, out int rowIndex, out int columnIndex)
        {
            Parse(Normalize(address, operation), out var rowNumber, out var columnNumber);
            rowIndex = checked((int)rowNumber - 1);
            columnIndex = columnNumber - 1;
        }

        internal static void Parse(string normalizedAddress, out uint rowNumber, out int columnNumber)
        {
            var separator = normalizedAddress.TakeWhile(character => character >= 'A' && character <= 'Z').Count();
            rowNumber = uint.Parse(normalizedAddress.Substring(separator), CultureInfo.InvariantCulture);
            columnNumber = GetColumnNumber(normalizedAddress.Substring(0, separator));
        }

        internal static int GetColumnNumber(string? address)
        {
            if (string.IsNullOrEmpty(address))
            {
                return int.MaxValue;
            }

            var number = 0;
            foreach (var character in address!)
            {
                if (character < 'A' || character > 'Z')
                {
                    break;
                }

                number = checked(number * 26 + character - 'A' + 1);
            }

            return number;
        }

        internal static int ValidateRange(int startRowIndex, int startColumnIndex, int rowCount, int columnCount, string operation)
        {
            if (rowCount <= 0 || columnCount <= 0)
            {
                throw new ReportEngineException(
                    ReportErrorCode.InvalidRangeDimensions,
                    operation,
                    $"Range dimensions must be positive; received {rowCount} x {columnCount}.");
            }

            var lastRow = (long)startRowIndex + rowCount - 1;
            var lastColumn = (long)startColumnIndex + columnCount - 1;
            if (startRowIndex < 0 || startColumnIndex < 0 || lastRow > MaximumRowIndex || lastColumn > MaximumColumnIndex)
            {
                throw new ReportEngineException(
                    ReportErrorCode.InvalidCellAddress,
                    operation,
                    $"Range starting at ({startRowIndex}, {startColumnIndex}) with dimensions {rowCount} x {columnCount} exceeds Excel worksheet limits.");
            }

            try
            {
                return checked(rowCount * columnCount);
            }
            catch (OverflowException ex)
            {
                throw new ReportEngineException(ReportErrorCode.InvalidRangeDimensions, operation, "The requested range is too large.", ex);
            }
        }

        private static void ValidateCoordinates(int rowIndex, int columnIndex, string operation)
        {
            if (rowIndex < 0 || rowIndex > MaximumRowIndex || columnIndex < 0 || columnIndex > MaximumColumnIndex)
            {
                throw new ReportEngineException(
                    ReportErrorCode.InvalidCellAddress,
                    operation,
                    $"Cell coordinates ({rowIndex}, {columnIndex}) are outside the valid zero-based Excel range.");
            }
        }

        private static ReportEngineException Invalid(string? address, string operation)
        {
            return new ReportEngineException(ReportErrorCode.InvalidCellAddress, operation, $"Invalid A1 cell address: {address ?? string.Empty}");
        }
    }
}
