namespace excelReport.Models;

/// <summary>單一報表的完整設定，序列化存放於 App_Data/configs/{code}.json。</summary>
public class ReportConfig
{
    public string Code { get; set; } = "";
    public string Name { get; set; } = "";
    public string TemplateFile { get; set; } = "";
    /// <summary>要讀取範本檔的哪個工作表：數字視為 0 起始的順序索引，字串視為工作表名稱。
    /// 留空時 fallback 到 appsettings 的 TemplateScan:Sheet（沿用舊設定檔行為）。</summary>
    public string TemplateSheet { get; set; } = "";
    public DataSourceConfig DataSource { get; set; } = new();
    public List<ParameterDef> Parameters { get; set; } = new();
    public MappingConfig Mapping { get; set; } = new();
    /// <summary>子報表：獨立的 xlsx 範本檔，產生報表時依序渲染、接在主範本（以及前一個子報表）
    /// 後面。每個子報表各自綁定一組父層/子層陣列（如 tasks[].items[]），依父層元素數量重複整份
    /// 子範本；同一個 ReportConfig 可以設定多個子報表，依清單順序依序附加。</summary>
    public List<SubReportConfig> SubReports { get; set; } = new();
}

/// <summary>
/// 資料來源設定。Type=url（預設）時代打外部 API；Type=paste 時直接使用 StaticJson
/// 這段貼上的固定 JSON 內容，不會發送任何 HTTP 請求，方便在還沒有真實 API 時先行測試映射設定。
/// </summary>
public class DataSourceConfig
{
    public string Type { get; set; } = "url";
    public string Url { get; set; } = "";
    public string Method { get; set; } = "GET";
    public Dictionary<string, string> Headers { get; set; } = new();
    public int TimeoutSec { get; set; } = 30;
    public string StaticJson { get; set; } = "";
}

/// <summary>報表輸入參數定義，用於 /Report/Index 動態產生輸入欄位。</summary>
public class ParameterDef
{
    public string Key { get; set; } = "";
    public string Label { get; set; } = "";
    public string Type { get; set; } = "string";
    public bool Required { get; set; }
}

/// <summary>
/// 範本標記與 JSONPath 的映射設定。Fields/Collections.Columns 的 value 是 JSONPath 字串；
/// 留空時 ReportEngine 會自動用 "$.{標記名稱}" 當預設路徑，不需要每個標記都手動填寫。
/// </summary>
public class MappingConfig
{
    public string Root { get; set; } = "$";
    public Dictionary<string, string> Fields { get; set; } = new();
    public List<CollectionMapping> Collections { get; set; } = new();
    public Dictionary<string, ImageMapping> Images { get; set; } = new();
}

/// <summary>
/// 一份子報表：獨立的 xlsx 範本檔，綁定一組父層/子層陣列(如 tasks[].items[])，渲染時對父層
/// 陣列的每個元素各自渲染一次整份子範本(父層集合永遠只塞一筆讓表頭列剛好展開一次；子層集合
/// 塞該父層底下全部元素)，再把渲染結果整份接到目前輸出的工作表後面。因為 MiniExcel 一個範本
/// 區域只能由一個集合驅動展開列數，沒辦法讓「父層列數」跟「子層列數」同時各自獨立變動，所以
/// 拆成獨立檔案、逐一渲染再拼接，而不是用一般 Collections 那種單層攤平的路徑。
/// </summary>
public class SubReportConfig
{
    public string TemplateFile { get; set; } = "";
    /// <summary>要讀取這份子報表範本檔的哪個工作表：數字視為 0 起始的順序索引，字串視為工作表
    /// 名稱。留空時 fallback 到 appsettings 的 TemplateScan:Sheet。</summary>
    public string TemplateSheet { get; set; } = "";
    /// <summary>父層陣列的 JSONPath，相對 root，例如 "$.tasks"。</summary>
    public string ParentPath { get; set; } = "";
    /// <summary>對應範本標記前綴，例如 {{tasks.xxx}} 就填 "tasks"。</summary>
    public string ParentCollectionName { get; set; } = "";
    /// <summary>父層欄位的 JSONPath，相對每個父層元素自己；留空時預設 "$.{欄位名}"。</summary>
    public Dictionary<string, string> ParentColumns { get; set; } = new();
    /// <summary>子層陣列的 JSONPath，相對每個父層元素，例如 "$.items"。</summary>
    public string ChildPath { get; set; } = "";
    /// <summary>對應範本標記前綴，例如 {{items.xxx}} 就填 "items"。</summary>
    public string ChildCollectionName { get; set; } = "";
    /// <summary>子層欄位的 JSONPath，相對每個子層元素自己；留空時預設 "$.{欄位名}"。
    /// 欄位名為 "seq" 時是保留字，自動填入該元素在陣列中的序號(從 1 開始)，不會走 JSONPath。</summary>
    public Dictionary<string, string> ChildColumns { get; set; } = new();
    public SpreadConfig? ChildSpread { get; set; }
}

/// <summary>集合（明細）映射設定。</summary>
public class CollectionMapping
{
    public string Name { get; set; } = "";
    public string Path { get; set; } = "";
    public Dictionary<string, string> Columns { get; set; } = new();
    public SpreadConfig? Spread { get; set; }
}

/// <summary>集合內陣列橫向攤平設定，例如把 values[] 攤平成 v1~v13。</summary>
public class SpreadConfig
{
    public string From { get; set; } = "";
    public string Prefix { get; set; } = "v";
    public int Max { get; set; } = 13;
}

/// <summary>
/// 圖片欄位映射。path 以 "$" 開頭時視為 JSONPath（指向的值可以是圖片 URL 或 base64 內容）；
/// 不以 "$" 開頭則直接當成固定值使用（例如寫死的圖片 URL），不會嘗試用 JSONPath 解析。
/// sourceType=auto（預設）依內容自動判斷（http(s):// 開頭視為 URL，否則視為 base64）；
/// 也可以明確指定 url 或 base64。找不到節點或抓取/解碼失敗時，該圖片直接略過，不會中斷報表產生。
/// </summary>
public class ImageMapping
{
    public string Path { get; set; } = "";
    public string SourceType { get; set; } = "auto";
}
