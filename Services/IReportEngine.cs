using excelReport.Models;

namespace excelReport.Services;

/// <summary>
/// 報表產生核心邏輯：讀設定 → 打資料來源 → JSONPath 映射 → 集合攤平 → 套版產生 xlsx。
/// 獨立於 Controller 之外，方便日後整包搬移到其他專案。
/// </summary>
public interface IReportEngine
{
    Task<ReportGenerationResult> GenerateAsync(string code, IDictionary<string, string> parameters, CancellationToken cancellationToken = default);
}
