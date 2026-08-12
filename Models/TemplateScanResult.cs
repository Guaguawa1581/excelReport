namespace excelReport.Models;

/// <summary>範本標記掃描結果，供設定編輯頁顯示與映射驗證使用。</summary>
public class TemplateScanResult
{
    /// <summary>不含點號的一般欄位標記，例如 {{docNo}}。</summary>
    public List<string> Fields { get; set; } = new();

    /// <summary>含點號的集合標記，依集合名稱分組，例如 {{items.seq}}。</summary>
    public List<CollectionScanResult> Collections { get; set; } = new();

    /// <summary>圖片標記，例如 [[photo]]（方括號語法，與一般文字標記的 {{}} 語法分開，
    /// 避免 MiniExcel 套版時誤處理——MiniExcel 完全不認得 [[]]，等它套版完再由
    /// ReportEngine 另外掃描、下載/解碼並嵌入圖片）。</summary>
    public List<string> ImageFields { get; set; } = new();
}

/// <summary>單一集合標記分組結果。</summary>
public class CollectionScanResult
{
    public string Name { get; set; } = "";
    public List<string> Columns { get; set; } = new();
}
