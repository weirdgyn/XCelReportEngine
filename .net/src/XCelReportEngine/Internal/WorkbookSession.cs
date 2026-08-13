using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace XCelReportEngine.Internal
{
    internal sealed class WorkbookSession : IDisposable
    {
        private bool _disposed;

        internal WorkbookSession(int id, string sourcePath, string outputPath, string stagingPath, SpreadsheetDocument document)
        {
            Id = id;
            SourcePath = sourcePath;
            OutputPath = outputPath;
            StagingPath = stagingPath;
            Document = document;

            var firstSheet = (Workbook.Sheets?.Elements<Sheet>().FirstOrDefault()) ?? throw new ReportEngineException(ReportErrorCode.InvalidWorkbook, "OpenWorkbook", "The workbook does not contain any worksheets.");
            ActiveSheetId = firstSheet.Id?.Value ?? string.Empty;
        }

        internal int Id { get; }

        internal string SourcePath { get; }

        internal string OutputPath { get; }

        internal string StagingPath { get; }

        internal SpreadsheetDocument Document { get; }

        internal WorkbookPart WorkbookPart => Document.WorkbookPart
            ?? throw new ReportEngineException(ReportErrorCode.InvalidWorkbook, "WorkbookPart", "The workbook part is missing.");

        internal Workbook Workbook => WorkbookPart.Workbook
            ?? throw new ReportEngineException(ReportErrorCode.InvalidWorkbook, "Workbook", "The workbook root element is missing.");

        internal string ActiveSheetId { get; set; }

        internal object SyncRoot { get; } = new object();

        internal Sheet ActiveSheet
        {
            get
            {
                EnsureOpen();
                var sheet = Workbook.Sheets?.Elements<Sheet>()
                    .FirstOrDefault(item => string.Equals(item.Id?.Value, ActiveSheetId, StringComparison.Ordinal));

                return sheet ?? throw new ReportEngineException(
                    ReportErrorCode.WorksheetNotFound,
                    "ActiveWorksheet",
                    "The active worksheet no longer exists.");
            }
        }

        internal WorksheetPart ActiveWorksheetPart
        {
            get
            {
                var relationshipId = ActiveSheet.Id?.Value;
                if (string.IsNullOrEmpty(relationshipId))
                {
                    throw new ReportEngineException(ReportErrorCode.InvalidWorkbook, "ActiveWorksheet", "The worksheet relationship is missing.");
                }

                return WorkbookPart.GetPartById(relationshipId!) as WorksheetPart
                    ?? throw new ReportEngineException(ReportErrorCode.InvalidWorkbook, "ActiveWorksheet", "The worksheet part is missing.");
            }
        }

        internal void EnsureOpen()
        {
            if (_disposed)
            {
                throw new ReportEngineException(ReportErrorCode.SessionClosed, "Session", $"Report session {Id} is closed.");
            }
        }

        internal void SaveToOutput()
        {
            EnsureOpen();
            Workbook.Save();
            Document.Save();

            var outputDirectory = Path.GetDirectoryName(OutputPath);
            if (string.IsNullOrEmpty(outputDirectory) || !Directory.Exists(outputDirectory))
            {
                throw new ReportEngineException(ReportErrorCode.OutputDirectoryNotFound, "SaveWorkbook", $"Output directory not found: {outputDirectory}");
            }

            var publishPath = Path.Combine(outputDirectory, $".{Path.GetFileName(OutputPath)}.{Guid.NewGuid():N}.tmp");
            try
            {
                using (var clone = Document.Clone(publishPath, true))
                {
                    clone.Save();
                }

                if (File.Exists(OutputPath))
                {
                    File.Replace(publishPath, OutputPath, null, true);
                }
                else
                {
                    File.Move(publishPath, OutputPath);
                }
            }
            catch (IOException ex)
            {
                TryDelete(publishPath);
                throw new ReportEngineException(ReportErrorCode.OutputFileLocked, "SaveWorkbook", $"Unable to publish workbook: {OutputPath}", ex);
            }
            catch (UnauthorizedAccessException ex)
            {
                TryDelete(publishPath);
                throw new ReportEngineException(ReportErrorCode.OutputFileLocked, "SaveWorkbook", $"Unable to publish workbook: {OutputPath}", ex);
            }
            catch (Exception ex) when (ex is not ReportEngineException)
            {
                TryDelete(publishPath);
                throw new ReportEngineException(ReportErrorCode.SaveFailed, "SaveWorkbook", $"Unable to save workbook: {OutputPath}", ex);
            }
        }

        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }

            _disposed = true;
            Document.Dispose();
            TryDelete(StagingPath);
        }

        private static void TryDelete(string path)
        {
            if (string.IsNullOrEmpty(path) || !File.Exists(path))
            {
                return;
            }

            try
            {
                File.Delete(path);
            }
            catch
            {
                // Cleanup is best effort; the primary operation error is more useful to LabVIEW.
            }
        }
    }
}
