namespace excelReport.Models;

/// <summary>單一報表的完整設定，序列化存放於 App_Data/configs/{logId}.json。</summary>
public class ReportConfig
{
    /// <summary>永久唯一識別碼（uuid）：建立時產生一次，之後不會改變。設定檔的檔名、以及儲存層
    /// 內部查找都以 LogId 為準，Code 只是給人看、給網址用的識別碼，因此可以事後自由改名而不影響
    /// 既有設定的存取。留空時 ReportConfigStore 儲存時會自動補上一個新的 uuid。</summary>
    public string LogId { get; set; } = "";
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
    /// 後面。每個子報表各自綁定一個根節點 JSONPath，查到幾個節點就重複渲染幾次整份子範本；
    /// 同一個 ReportConfig 可以設定多個子報表，依清單順序依序附加。</summary>
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
/// 一份子報表：獨立的 xlsx 範本檔，綁定一個根節點 JSONPath（相對主報表 root）。用標準 JSONPath
/// 語法解析（SelectTokens）：查到幾個節點，就對每個節點各自渲染一次整份子範本、依序接到目前
/// 輸出的工作表後面。帶萬用字元的路徑（例如 "$.tasks[*]"）語意上查到多個節點，就是迴圈報表；
/// 不帶萬用字元的路徑（例如 "$"、"$.data"、"$.records[0]"）語意上剛好查到一個節點，就只渲染
/// 一次（取特定值，不迴圈）。每次渲染時，Mapping（Fields/Collections/Images）都相對「這次查到
/// 的節點」解析，用法跟主報表 MappingConfig 完全一致。
/// </summary>
public class SubReportConfig
{
    public string TemplateFile { get; set; } = "";
    /// <summary>要讀取這份子報表範本檔的哪個工作表：數字視為 0 起始的順序索引，字串視為工作表
    /// 名稱。留空時 fallback 到 appsettings 的 TemplateScan:Sheet。</summary>
    public string TemplateSheet { get; set; } = "";
    /// <summary>根節點 JSONPath，相對主報表 root，用標準 JSONPath 語法解析（SelectTokens）。
    /// 查到幾個節點就渲染幾次：帶萬用字元的路徑（例如 "$.tasks[*]"）查到多個節點就是迴圈報表；
    /// 不帶萬用字元的路徑（例如 "$"、"$.data"、"$.records[0]"）查到單一節點就只渲染一次。</summary>
    public string Root { get; set; } = "";
    /// <summary>子報表本身的欄位映射，用法跟主報表 MappingConfig 完全一致：Fields 是相對每次
    /// 渲染當下節點的單值欄位（{{field}}）、Collections 是集合明細（{{collection.col}}，可以有
    /// 多個、各自可搭配 Spread 攤平設定）、Images 是圖片標記（[[field]]）。Mapping.Root 對子報表
    /// 沒有意義，固定忽略不使用。</summary>
    public MappingConfig Mapping { get; set; } = new();
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
    /// <summary>來源陣列的 JSONPath，相對每筆項目，例如 "$.samples"。</summary>
    public string From { get; set; } = "";
    /// <summary>From 陣列的元素若為物件，指定要取用的屬性，用標準 JSONPath 語法解析（SelectToken），
    /// root 是 From（spreadFrom）查到的每個元素本身，例如 "$.result"（也接受省略開頭 "$." 的簡寫，
    /// 例如 "result"，語意相同）；留空時直接使用元素本身的值(純量陣列，或 From 本來就用萬用字元/
    /// 點記法指到純量，例如 "$.samples[*].result")。</summary>
    public string? Property { get; set; }
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
