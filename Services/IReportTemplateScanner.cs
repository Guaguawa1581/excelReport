using DocumentFormat.OpenXml.Packaging;
using excelReport.Models;

namespace excelReport.Services;

/// <summary>掃描 xlsx 範本內所有 {{xxx}} 標記，分類成一般欄位與集合欄位。</summary>
public interface IReportTemplateScanner
{
    TemplateScanResult Scan(byte[] templateBytes);

    /// <summary>回傳只保留目標工作表（TemplateScan:Sheet 設定指定的那一個）的範本內容，
    /// 其餘工作表整個移除。範本只有一個工作表時原樣回傳。僅供套版前暫用，
    /// 套版結果需搭配 <see cref="MergeRenderedTargetSheet"/> 貼回完整範本。</summary>
    byte[] IsolateTargetSheet(byte[] templateBytes);

    /// <summary>把只含目標工作表的套版結果（<paramref name="renderedTargetSheetBytes"/>，來自
    /// IsolateTargetSheet 輸出再經 MiniExcel 套版）貼回完整的原始範本，取代其中目標工作表的內容，
    /// 其餘工作表維持原樣（位元組層級不變）。</summary>
    byte[] MergeRenderedTargetSheet(byte[] originalTemplateBytes, byte[] renderedTargetSheetBytes);

    /// <summary>依 TemplateScan:Sheet 設定解析出目標工作表的 WorksheetPart。</summary>
    WorksheetPart GetTargetWorksheetPart(WorkbookPart workbookPart);
}
