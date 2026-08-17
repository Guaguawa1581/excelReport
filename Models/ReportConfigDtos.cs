namespace excelReport.Models;

/// <summary>上傳/掃描範本後回傳的封裝：掃描結果對應目前選定的那個工作表；SheetNames 是整份
/// 檔案依原始順序排列的所有工作表名稱，供畫面畫出「(索引)名稱」選單讓使用者換頁重新掃描。</summary>
public class UploadTemplateResult
{
    public string FileName { get; set; } = "";
    public TemplateScanResult ScanResult { get; set; } = new();
    public List<string> SheetNames { get; set; } = new();
}

/// <summary>「測試連線」AJAX 請求內容。</summary>
public class TestConnectionRequest
{
    public string Type { get; set; } = "url";
    public string Url { get; set; } = "";
    public string Method { get; set; } = "GET";
    public Dictionary<string, string> Headers { get; set; } = new();
    public Dictionary<string, string> Parameters { get; set; } = new();
    public int TimeoutSec { get; set; } = 30;
    public string StaticJson { get; set; } = "";
}

/// <summary>「測試連線」AJAX 回應內容。</summary>
public class TestConnectionResponse
{
    public bool Success { get; set; }
    public string? Body { get; set; }
    public string? ErrorMessage { get; set; }
}
