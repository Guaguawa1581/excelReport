using excelReport.Models;

namespace excelReport.Services;

/// <summary>報表設定的存取介面，底層以 App_Data/configs/{code}.json 存放。</summary>
public interface IReportConfigStore
{
    List<ReportConfig> GetAll();
    ReportConfig? Get(string code);
    void Save(ReportConfig config);
    void Delete(string code);
    bool Exists(string code);
}
