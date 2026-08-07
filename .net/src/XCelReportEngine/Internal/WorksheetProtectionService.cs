using System;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Spreadsheet;

namespace XCelReportEngine.Internal
{
    internal static class WorksheetProtectionService
    {
        internal static void LockAll(WorkbookSession session, string password)
        {
            if (password == null)
            {
                throw new ReportEngineException(ReportErrorCode.InvalidArgument, "LockAllWorksheets", "Password cannot be null.");
            }

            var hash = ConvertPasswordToLegacyHash(password);
            foreach (var part in session.WorkbookPart.WorksheetParts)
            {
                var worksheet = part.Worksheet
                    ?? throw new ReportEngineException(ReportErrorCode.InvalidWorkbook, "LockAllWorksheets", "Worksheet root element is missing.");
                worksheet.RemoveAllChildren<SheetProtection>();

                var protection = new SheetProtection
                {
                    Sheet = true,
                    Objects = true,
                    Scenarios = true,
                    Password = hash
                };

                var calculationProperties = worksheet.Elements<SheetCalculationProperties>().LastOrDefault();
                if (calculationProperties != null)
                {
                    worksheet.InsertAfter(protection, calculationProperties);
                }
                else
                {
                    var sheetData = worksheet.GetFirstChild<SheetData>()
                        ?? throw new ReportEngineException(ReportErrorCode.InvalidWorkbook, "LockAllWorksheets", "Worksheet is missing SheetData.");
                    worksheet.InsertAfter(protection, sheetData);
                }

                worksheet.Save();
            }
        }

        internal static void UnlockAll(WorkbookSession session)
        {
            foreach (var part in session.WorkbookPart.WorksheetParts)
            {
                var worksheet = part.Worksheet
                    ?? throw new ReportEngineException(ReportErrorCode.InvalidWorkbook, "UnlockAllWorksheets", "Worksheet root element is missing.");
                worksheet.RemoveAllChildren<SheetProtection>();
                worksheet.Save();
            }
        }

        internal static string ConvertPasswordToLegacyHash(string password)
        {
            var characters = Encoding.ASCII.GetBytes(password);
            var hash = 0;

            if (characters.Length > 0)
            {
                var index = characters.Length;
                while (index-- > 0)
                {
                    hash = ((hash >> 14) & 0x01) | ((hash << 1) & 0x7fff);
                    hash ^= characters[index];
                }

                hash = ((hash >> 14) & 0x01) | ((hash << 1) & 0x7fff);
                hash ^= characters.Length;
                hash ^= 0x8000 | ('N' << 8) | 'K';
            }

            return Convert.ToString(hash, 16).ToUpperInvariant();
        }
    }
}
