using System;
using System.Linq;
using DocumentFormat.OpenXml.Spreadsheet;
using XCelReportEngine.Internal;

namespace XCelReportEngine
{
    public sealed class ReportEngineApi : IDisposable
    {
        private readonly SessionRegistry _sessions = new SessionRegistry();
        private bool _disposed;

        public static ReportEngineApi Create()
        {
            return new ReportEngineApi();
        }

        public static string ConvertCellIndexToAddress(int rowIndex, int columnIndex)
        {
            return CellAddress.FromZeroBased(rowIndex, columnIndex, "ConvertCellIndexToAddress");
        }

        public static int[] ConvertCellAddressToIndex(string cellAddress)
        {
            CellAddress.ToCoordinates(
                cellAddress,
                "ConvertCellAddressToIndex",
                out var rowIndex,
                out var columnIndex);
            return new[] { rowIndex, columnIndex };
        }

        public int OpenWorkbook(string sourcePath, string outputPath)
        {
            EnsureOpen();
            return _sessions.Open(sourcePath, outputPath);
        }

        public string[] GetWorksheetNames(int sessionId)
        {
            var session = GetSession(sessionId);
            lock (session.SyncRoot)
            {
                return session.Workbook.Sheets?.Elements<Sheet>()
                    .Select(item => item.Name?.Value ?? string.Empty)
                    .ToArray() ?? Array.Empty<string>();
            }
        }

        public string GetActiveWorksheetName(int sessionId)
        {
            var session = GetSession(sessionId);
            lock (session.SyncRoot)
            {
                return session.ActiveSheet.Name?.Value ?? string.Empty;
            }
        }

        public void SelectWorksheetByName(int sessionId, string worksheetName)
        {
            if (string.IsNullOrWhiteSpace(worksheetName))
            {
                throw new ReportEngineException(ReportErrorCode.InvalidArgument, "SelectWorksheetByName", "Worksheet name is required.");
            }

            var session = GetSession(sessionId);
            lock (session.SyncRoot)
            {
                var sheet = session.Workbook.Sheets?.Elements<Sheet>()
                    .FirstOrDefault(item => string.Equals(item.Name?.Value, worksheetName, StringComparison.Ordinal));
                if (sheet == null)
                {
                    throw new ReportEngineException(ReportErrorCode.WorksheetNotFound, "SelectWorksheetByName", $"Worksheet not found: {worksheetName}");
                }

                session.ActiveSheetId = sheet.Id?.Value ?? string.Empty;
            }
        }

        public void SelectWorksheetByIndex(int sessionId, int worksheetIndex)
        {
            var session = GetSession(sessionId);
            lock (session.SyncRoot)
            {
                var sheets = session.Workbook.Sheets?.Elements<Sheet>().ToArray() ?? Array.Empty<Sheet>();
                if (worksheetIndex < 0 || worksheetIndex >= sheets.Length)
                {
                    throw new ReportEngineException(
                        ReportErrorCode.WorksheetIndexOutOfRange,
                        "SelectWorksheetByIndex",
                        $"Worksheet index {worksheetIndex} is outside the valid range 0..{sheets.Length - 1}.");
                }

                session.ActiveSheetId = sheets[worksheetIndex].Id?.Value ?? string.Empty;
            }
        }

        public void LockAllWorksheets(int sessionId, string password)
        {
            var session = GetSession(sessionId);
            lock (session.SyncRoot)
            {
                WorksheetProtectionService.LockAll(session, password);
            }
        }

        public void UnlockAllWorksheets(int sessionId)
        {
            var session = GetSession(sessionId);
            lock (session.SyncRoot)
            {
                WorksheetProtectionService.UnlockAll(session);
            }
        }

        public string ReadCellStringByAddress(int sessionId, string cellAddress)
        {
            const string operation = "ReadCellStringByAddress";
            var session = GetSession(sessionId);
            lock (session.SyncRoot)
            {
                return CellService.ReadAsString(session, cellAddress, operation);
            }
        }

        public string ReadCellStringByIndex(int sessionId, int rowIndex, int columnIndex)
        {
            const string operation = "ReadCellStringByIndex";
            var session = GetSession(sessionId);
            lock (session.SyncRoot)
            {
                var address = CellAddress.FromZeroBased(rowIndex, columnIndex, operation);
                return CellService.ReadAsString(session, address, operation);
            }
        }

        public void WriteCellStringByAddress(int sessionId, string cellAddress, string value)
        {
            const string operation = "WriteCellStringByAddress";
            var session = GetSession(sessionId);
            lock (session.SyncRoot)
            {
                CellService.WriteString(session, cellAddress, value, operation);
            }
        }

        public void WriteCellStringByIndex(int sessionId, int rowIndex, int columnIndex, string value)
        {
            const string operation = "WriteCellStringByIndex";
            var session = GetSession(sessionId);
            lock (session.SyncRoot)
            {
                CellService.WriteString(session, CellAddress.FromZeroBased(rowIndex, columnIndex, operation), value, operation);
            }
        }

        public void WriteCellDoubleByAddress(int sessionId, string cellAddress, double value)
        {
            const string operation = "WriteCellDoubleByAddress";
            var session = GetSession(sessionId);
            lock (session.SyncRoot)
            {
                CellService.WriteDouble(session, cellAddress, value, operation);
            }
        }

        public void WriteCellDoubleByIndex(int sessionId, int rowIndex, int columnIndex, double value)
        {
            const string operation = "WriteCellDoubleByIndex";
            var session = GetSession(sessionId);
            lock (session.SyncRoot)
            {
                CellService.WriteDouble(session, CellAddress.FromZeroBased(rowIndex, columnIndex, operation), value, operation);
            }
        }

        public void WriteCellBooleanByAddress(int sessionId, string cellAddress, bool value)
        {
            const string operation = "WriteCellBooleanByAddress";
            var session = GetSession(sessionId);
            lock (session.SyncRoot)
            {
                CellService.WriteBoolean(session, cellAddress, value, operation);
            }
        }

        public void WriteCellBooleanByIndex(int sessionId, int rowIndex, int columnIndex, bool value)
        {
            const string operation = "WriteCellBooleanByIndex";
            var session = GetSession(sessionId);
            lock (session.SyncRoot)
            {
                CellService.WriteBoolean(session, CellAddress.FromZeroBased(rowIndex, columnIndex, operation), value, operation);
            }
        }

        public string[] ReadStringRangeByIndex(
            int sessionId,
            int startRowIndex,
            int startColumnIndex,
            int rowCount,
            int columnCount)
        {
            const string operation = "ReadStringRangeByIndex";
            var session = GetSession(sessionId);
            lock (session.SyncRoot)
            {
                return CellService.ReadStringRange(
                    session,
                    startRowIndex,
                    startColumnIndex,
                    rowCount,
                    columnCount,
                    operation);
            }
        }

        public void WriteStringRangeByIndex(
            int sessionId,
            int startRowIndex,
            int startColumnIndex,
            int rowCount,
            int columnCount,
            string[] values)
        {
            const string operation = "WriteStringRangeByIndex";
            var session = GetSession(sessionId);
            lock (session.SyncRoot)
            {
                CellService.WriteStringRange(
                    session,
                    startRowIndex,
                    startColumnIndex,
                    rowCount,
                    columnCount,
                    values,
                    operation);
            }
        }

        public int AppendImage(
            int sessionId,
            string imagePath,
            int rowIndex,
            int columnIndex,
            string cellAddress,
            int alignment,
            string caption)
        {
            const string operation = "AppendImage";
            if (alignment < 0 || alignment > 8)
            {
                throw new ReportEngineException(ReportErrorCode.InvalidArgument, operation, $"Invalid image alignment value: {alignment}");
            }

            var session = GetSession(sessionId);
            lock (session.SyncRoot)
            {
                // Excel ignores the alignment and caption inputs of the NI VI. They are retained in this
                // signature so the LabVIEW wrapper can remain connector-compatible.
                _ = caption;
                return ImageService.Append(session, imagePath, rowIndex, columnIndex, cellAddress, operation);
            }
        }

        public void FormatImage(
            int sessionId,
            int measurementSystem,
            double scaleFactor,
            int pictureIndex,
            double height,
            double width,
            int colorType)
        {
            const string operation = "FormatImage";
            var session = GetSession(sessionId);
            lock (session.SyncRoot)
            {
                ImageService.Format(session, measurementSystem, scaleFactor, pictureIndex, height, width, colorType, operation);
            }
        }

        public void SaveWorkbook(int sessionId)
        {
            var session = GetSession(sessionId);
            lock (session.SyncRoot)
            {
                session.SaveToOutput();
            }
        }

        public void CloseWorkbook(int sessionId, bool saveChanges)
        {
            EnsureOpen();
            _sessions.Close(sessionId, saveChanges);
        }

        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }

            _disposed = true;
            _sessions.Dispose();
        }

        private WorkbookSession GetSession(int sessionId)
        {
            EnsureOpen();
            return _sessions.Get(sessionId);
        }

        private void EnsureOpen()
        {
            if (_disposed)
            {
                throw new ObjectDisposedException(nameof(ReportEngineApi));
            }
        }
    }
}
