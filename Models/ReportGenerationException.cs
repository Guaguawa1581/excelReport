namespace excelReport.Models;

/// <summary>報表產生流程中已知的錯誤，訊息為可直接顯示給使用者的繁體中文說明。</summary>
public class ReportGenerationException : Exception
{
    public ReportGenerationException(string message) : base(message) { }

    public ReportGenerationException(string message, Exception innerException) : base(message, innerException) { }
}
