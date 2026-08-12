namespace excelReport.Models;

/// <summary>/Report/Index 頁面用的清單項目。</summary>
public class ReportListItemViewModel
{
    public string Code { get; set; } = "";
    public string Name { get; set; } = "";
    public List<ParameterDef> Parameters { get; set; } = new();
    public string TemplateFile { get; set; } = "";
}

/// <summary>/ReportConfig/Edit 頁面的編輯用視圖模型。</summary>
public class ReportConfigEditViewModel
{
    public bool IsNew { get; set; }
    public string OriginalCode { get; set; } = "";
    public ReportConfig Config { get; set; } = new();

    /// <summary>目前範本檔掃描出的標記結果（若有上傳/已存在範本）。</summary>
    public TemplateScanResult? ScanResult { get; set; }

    /// <summary>可用的範本檔清單（App_Data/templates 下的 .xlsx 檔名）。</summary>
    public List<string> AvailableTemplates { get; set; } = new();

    /// <summary>測試連線後端回傳的原始 JSON 內容。</summary>
    public string? TestConnectionResult { get; set; }
    public bool TestConnectionSuccess { get; set; }

    public string? ErrorMessage { get; set; }
    public string? SuccessMessage { get; set; }
}

/// <summary>錯誤顯示頁面用的視圖模型。</summary>
public class ReportErrorViewModel
{
    public string Title { get; set; } = "發生錯誤";
    public string Message { get; set; } = "";
    public string? Detail { get; set; }
}
