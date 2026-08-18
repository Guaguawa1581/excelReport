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

    /// <summary>範本裡實際出現過的 {{__end__}} 大括號內文字，含使用者打的任何空白變體（例如
    /// "__end__" 或 " __end__ "）。MiniExcel 自己的套版引擎是直接拿大括號內的原始文字（不會像
    /// 我們自己的掃描正則那樣先 trim 掉空白）當資料字典的 key 查值，所以要把每一種實際出現過的
    /// 原始寫法都當成 key 補值，範本作者不管有沒有在大括號內加空白都能正常運作。</summary>
    public List<string> EndMarkerRawKeys { get; set; } = new();
}

/// <summary>單一集合標記分組結果。</summary>
public class CollectionScanResult
{
    public string Name { get; set; } = "";
    public List<string> Columns { get; set; } = new();
}
