using excelReport.Models;

namespace excelReport.Services;

/// <summary>依報表設定的資料來源定義，代打外部 API 取回原始 JSON 字串。</summary>
public interface IDataSourceClient
{
    Task<string> FetchAsync(DataSourceConfig dataSource, IDictionary<string, string> parameters, CancellationToken cancellationToken = default);
}
