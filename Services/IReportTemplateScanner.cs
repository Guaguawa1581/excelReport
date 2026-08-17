using DocumentFormat.OpenXml.Packaging;
using excelReport.Models;

namespace excelReport.Services;

/// <summary>掃描 xlsx 範本內所有 {{xxx}} 標記，分類成一般欄位與集合欄位。</summary>
public interface IReportTemplateScanner
{
    /// <summary>sheet 是每份範本各自存的工作表選擇(數字視為 0 起始的順序索引，字串視為工作表
    /// 名稱)；不填時 fallback 到 appsettings 的 TemplateScan:Sheet。</summary>
    TemplateScanResult Scan(byte[] templateBytes, string? sheet = null);

    /// <summary>範本是否為「嚴格開放的 XML 試算表」(Strict Open XML Spreadsheet) 格式。
    /// MiniExcel 套版引擎不支援這個格式，即使掃描能成功，套版時也一定會丟出 NullReferenceException。</summary>
    bool IsStrictOpenXml(byte[] templateBytes);

    /// <summary>依 sheet 參數(或 fallback 到 appsettings 的 TemplateScan:Sheet)決定要處理哪個
    /// 工作表(數字視為 0 起始的順序索引，字串視為工作表名稱)。Scan 用它決定要掃描哪頁；
    /// ReportEngine 渲染巢狀子表區塊時也用同一個方法，確保取到的是同一個工作表。</summary>
    WorksheetPart GetTargetWorksheetPart(WorkbookPart workbookPart, string? sheet = null);

    /// <summary>依原始 xlsx 內的排列順序回傳所有工作表名稱，供設定頁面畫出選單讓使用者選擇。</summary>
    List<string> GetSheetNames(byte[] templateBytes);
}
