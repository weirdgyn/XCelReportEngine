namespace XCelReportEngine
{
    public enum ReportErrorCode
    {
        Unknown = 0,
        InvalidArgument = 1000,
        FileNotFound = 1001,
        UnsupportedFileType = 1002,
        OutputDirectoryNotFound = 1003,
        OutputFileLocked = 1004,
        InvalidWorkbook = 1100,
        WorksheetNotFound = 1101,
        WorksheetIndexOutOfRange = 1102,
        InvalidCellAddress = 1103,
        FormulaReadNotSupported = 1104,
        InvalidRangeDimensions = 1105,
        ImageFileNotFound = 1110,
        UnsupportedImageType = 1111,
        InvalidImage = 1112,
        PictureIndexOutOfRange = 1113,
        UnsupportedPictureColorType = 1114,
        SessionNotFound = 1200,
        SessionClosed = 1201,
        SaveFailed = 1300
    }
}
