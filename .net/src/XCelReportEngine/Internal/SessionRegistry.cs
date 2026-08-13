using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

namespace XCelReportEngine.Internal
{
    internal sealed class SessionRegistry : IDisposable
    {
        private readonly object _syncRoot = new();
        private readonly Dictionary<int, WorkbookSession> _sessions = [];
        private int _nextId;
        private bool _disposed;

        internal int Open(string sourcePath, string outputPath)
        {
            EnsureRegistryOpen();
            ValidatePaths(sourcePath, outputPath);

            var sourceFullPath = Path.GetFullPath(sourcePath);
            var outputFullPath = Path.GetFullPath(outputPath);
            var extension = Path.GetExtension(sourceFullPath);
            var isTemplate = string.Equals(extension, ".xltx", StringComparison.OrdinalIgnoreCase);
            var outputDirectory = Path.GetDirectoryName(outputFullPath) ?? string.Empty;
            var stagingPath = Path.Combine(outputDirectory, $".{Path.GetFileName(outputFullPath)}.{Guid.NewGuid():N}.working.tmp");

            SpreadsheetDocument? document = null;
            try
            {
                File.Copy(sourceFullPath, stagingPath, false);
                document = SpreadsheetDocument.Open(stagingPath, true);
                if (isTemplate)
                {
                    document.ChangeDocumentType(SpreadsheetDocumentType.Workbook);
                    document.Save();
                }

                var id = Interlocked.Increment(ref _nextId);
                var session = new WorkbookSession(id, sourceFullPath, outputFullPath, stagingPath, document);
                document = null;

                lock (_syncRoot)
                {
                    _sessions.Add(id, session);
                }

                return id;
            }
            catch (ReportEngineException)
            {
                document?.Dispose();
                TryDelete(stagingPath);
                throw;
            }
            catch (Exception ex)
            {
                document?.Dispose();
                TryDelete(stagingPath);
                throw new ReportEngineException(ReportErrorCode.InvalidWorkbook, "OpenWorkbook", $"Unable to open workbook: {sourceFullPath}", ex);
            }
        }

        internal WorkbookSession Get(int sessionId)
        {
            EnsureRegistryOpen();
            lock (_syncRoot)
            {
                if (_sessions.TryGetValue(sessionId, out var session))
                {
                    return session;
                }
            }

            throw new ReportEngineException(ReportErrorCode.SessionNotFound, "Session", $"Report session {sessionId} was not found.");
        }

        internal void Close(int sessionId, bool saveChanges)
        {
            WorkbookSession session;
            lock (_syncRoot)
            {
                if (!_sessions.TryGetValue(sessionId, out session!))
                {
                    throw new ReportEngineException(ReportErrorCode.SessionNotFound, "CloseWorkbook", $"Report session {sessionId} was not found.");
                }
            }

            lock (session.SyncRoot)
            {
                if (saveChanges)
                {
                    session.SaveToOutput();
                }

                session.Dispose();
            }

            lock (_syncRoot)
            {
                _sessions.Remove(sessionId);
            }
        }

        public void Dispose()
        {
            List<WorkbookSession> sessions;
            lock (_syncRoot)
            {
                if (_disposed)
                {
                    return;
                }

                _disposed = true;
                sessions = [.. _sessions.Values];
                _sessions.Clear();
            }

            foreach (var session in sessions)
            {
                lock (session.SyncRoot)
                {
                    session.Dispose();
                }
            }
        }

        private static void ValidatePaths(string sourcePath, string outputPath)
        {
            if (string.IsNullOrWhiteSpace(sourcePath))
            {
                throw new ReportEngineException(ReportErrorCode.InvalidArgument, "OpenWorkbook", "Source path is required.");
            }

            if (string.IsNullOrWhiteSpace(outputPath))
            {
                throw new ReportEngineException(ReportErrorCode.InvalidArgument, "OpenWorkbook", "Output path is required.");
            }

            if (!File.Exists(sourcePath))
            {
                throw new ReportEngineException(ReportErrorCode.FileNotFound, "OpenWorkbook", $"Source workbook not found: {sourcePath}");
            }

            var sourceExtension = Path.GetExtension(sourcePath);
            if (!string.Equals(sourceExtension, ".xlsx", StringComparison.OrdinalIgnoreCase)
                && !string.Equals(sourceExtension, ".xltx", StringComparison.OrdinalIgnoreCase))
            {
                throw new ReportEngineException(ReportErrorCode.UnsupportedFileType, "OpenWorkbook", $"Unsupported source type: {sourceExtension}");
            }

            if (!string.Equals(Path.GetExtension(outputPath), ".xlsx", StringComparison.OrdinalIgnoreCase))
            {
                throw new ReportEngineException(ReportErrorCode.UnsupportedFileType, "OpenWorkbook", "Output path must use the .xlsx extension.");
            }

            var outputDirectory = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (string.IsNullOrEmpty(outputDirectory) || !Directory.Exists(outputDirectory))
            {
                throw new ReportEngineException(ReportErrorCode.OutputDirectoryNotFound, "OpenWorkbook", $"Output directory not found: {outputDirectory}");
            }
        }

        private void EnsureRegistryOpen()
        {
            if (_disposed)
            {
                throw new ObjectDisposedException(nameof(SessionRegistry));
            }
        }

        private static void TryDelete(string path)
        {
            try
            {
                if (File.Exists(path))
                {
                    File.Delete(path);
                }
            }
            catch
            {
                // Best effort cleanup after a failed open.
            }
        }
    }
}
