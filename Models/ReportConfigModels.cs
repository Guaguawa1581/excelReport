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

/// <summary>資料來源 API 設定。</summary>
public class DataSourceConfig
{
    public string Url { get; set; } = "";
    public string Method { get; set; } = "GET";
    public Dictionary<string, string> Headers { get; set; } = new();
    public int TimeoutSec { get; set; } = 30;
}

/// <summary>報表輸入參數定義，用於 /Report/Index 動態產生輸入欄位。</summary>
public class ParameterDef
{
    public string Key { get; set; } = "";
    public string Label { get; set; } = "";
    public string Type { get; set; } = "string";
    public bool Required { get; set; }
}

/// <summary>範本標記與 JSONPath 的映射設定。</summary>
public class MappingConfig
{
    public string Root { get; set; } = "$";
    public Dictionary<string, string> Fields { get; set; } = new();
    public List<CollectionMapping> Collections { get; set; } = new();
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
