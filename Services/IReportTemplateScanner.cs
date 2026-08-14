using excelReport.Models;

namespace excelReport.Services;

/// <summary>掃描 xlsx 範本內所有 {{xxx}} 標記，分類成一般欄位與集合欄位。</summary>
public interface IReportTemplateScanner
{
    TemplateScanResult Scan(byte[] templateBytes);

    /// <summary>範本是否為「嚴格開放的 XML 試算表」(Strict Open XML Spreadsheet) 格式。
    /// MiniExcel 套版引擎不支援這個格式，即使掃描能成功，套版時也一定會丟出 NullReferenceException。</summary>
    bool IsStrictOpenXml(byte[] templateBytes);
}
