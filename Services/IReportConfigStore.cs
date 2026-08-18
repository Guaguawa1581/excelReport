using excelReport.Models;

namespace excelReport.Services;

/// <summary>報表設定的存取介面，底層以 App_Data/configs/{logId}.json 存放（LogId 是建立時產生
/// 的 uuid，不會改變）；Get/Delete/Exists 仍以 Code 查找，方便外部（網址、路由）沿用 Code 當識別碼。</summary>
public interface IReportConfigStore
{
    List<ReportConfig> GetAll();
    ReportConfig? Get(string code);
    void Save(ReportConfig config);
    void Delete(string code);
    bool Exists(string code);
}
