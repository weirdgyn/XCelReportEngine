using System;

namespace XCelReportEngine
{
    public sealed class ReportEngineException : Exception
    {
        public ReportEngineException(ReportErrorCode errorCode, string operation, string message)
            : base(message)
        {
            ErrorCode = (int)errorCode;
            ErrorCodeName = errorCode.ToString();
            Operation = operation ?? string.Empty;
        }

        public ReportEngineException(ReportErrorCode errorCode, string operation, string message, Exception innerException)
            : base(message, innerException)
        {
            ErrorCode = (int)errorCode;
            ErrorCodeName = errorCode.ToString();
            Operation = operation ?? string.Empty;
        }

        public int ErrorCode { get; }

        public string ErrorCodeName { get; }

        public string Operation { get; }
    }
}
