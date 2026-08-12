namespace excelReport.Models;

/// <summary>單一報表的完整設定，序列化存放於 App_Data/configs/{code}.json。</summary>
public class ReportConfig
{
    public string Code { get; set; } = "";
    public string Name { get; set; } = "";
    public string TemplateFile { get; set; } = "";
    public DataSourceConfig DataSource { get; set; } = new();
    public List<ParameterDef> Parameters { get; set; } = new();
    public MappingConfig Mapping { get; set; } = new();
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
