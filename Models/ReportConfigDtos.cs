namespace excelReport.Models;

/// <summary>上傳範本後回傳的掃描結果封裝。</summary>
public class UploadTemplateResult
{
    public string FileName { get; set; } = "";
    public TemplateScanResult ScanResult { get; set; } = new();
}

/// <summary>「測試連線」AJAX 請求內容。</summary>
public class TestConnectionRequest
{
    public string Url { get; set; } = "";
    public string Method { get; set; } = "GET";
    public Dictionary<string, string> Headers { get; set; } = new();
    public Dictionary<string, string> Parameters { get; set; } = new();
    public int TimeoutSec { get; set; } = 30;
}

/// <summary>「測試連線」AJAX 回應內容。</summary>
public class TestConnectionResponse
{
    public bool Success { get; set; }
    public string? Body { get; set; }
    public string? ErrorMessage { get; set; }
}
