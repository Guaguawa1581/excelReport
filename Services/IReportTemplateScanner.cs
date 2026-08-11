using excelReport.Models;

namespace excelReport.Services;

/// <summary>掃描 xlsx 範本內所有 {{xxx}} 標記，分類成一般欄位與集合欄位。</summary>
public interface IReportTemplateScanner
{
    TemplateScanResult Scan(byte[] templateBytes);
}
