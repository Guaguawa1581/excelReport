namespace excelReport.Models;

/// <summary>報表產生結果：檔案內容與建議下載檔名。</summary>
public class ReportGenerationResult
{
    public byte[] FileBytes { get; set; } = Array.Empty<byte>();
    public string FileName { get; set; } = "";
}
