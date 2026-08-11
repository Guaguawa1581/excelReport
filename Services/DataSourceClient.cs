using System.Net;
using excelReport.Models;

namespace excelReport.Services;

/// <summary>用 IHttpClientFactory 呼叫外部資料來源 API，統一處理逾時與非 2xx 錯誤。</summary>
public class DataSourceClient : IDataSourceClient
{
    private readonly IHttpClientFactory _httpClientFactory;

    public DataSourceClient(IHttpClientFactory httpClientFactory)
    {
        _httpClientFactory = httpClientFactory;
    }

    public async Task<string> FetchAsync(DataSourceConfig dataSource, IDictionary<string, string> parameters, CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrWhiteSpace(dataSource.Url))
        {
            throw new ReportGenerationException("資料來源 URL 未設定。");
        }

        var url = ApplyPlaceholders(dataSource.Url, parameters);

        using var client = _httpClientFactory.CreateClient("ReportDataSource");
        client.Timeout = TimeSpan.FromSeconds(dataSource.TimeoutSec > 0 ? dataSource.TimeoutSec : 30);

        var method = string.IsNullOrWhiteSpace(dataSource.Method) ? HttpMethod.Get : new HttpMethod(dataSource.Method.ToUpperInvariant());
        using var request = new HttpRequestMessage(method, url);
        foreach (var header in dataSource.Headers)
        {
            request.Headers.TryAddWithoutValidation(header.Key, header.Value);
        }

        using var timeoutCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        timeoutCts.CancelAfter(client.Timeout);

        HttpResponseMessage response;
        try
        {
            response = await client.SendAsync(request, timeoutCts.Token);
        }
        catch (OperationCanceledException) when (!cancellationToken.IsCancellationRequested)
        {
            throw new ReportGenerationException($"呼叫資料來源逾時（{dataSource.TimeoutSec} 秒），URL：{url}");
        }
        catch (HttpRequestException ex)
        {
            throw new ReportGenerationException($"無法連線至資料來源：{ex.Message}（URL：{url}）", ex);
        }

        using (response)
        {
            var body = await response.Content.ReadAsStringAsync(cancellationToken);
            if (!response.IsSuccessStatusCode)
            {
                throw new ReportGenerationException(
                    $"資料來源回應失敗，狀態碼 {(int)response.StatusCode} {response.StatusCode}。URL：{url}" +
                    (string.IsNullOrWhiteSpace(body) ? "" : $"\n回應內容：{Truncate(body, 500)}"));
            }

            if (string.IsNullOrWhiteSpace(body))
            {
                throw new ReportGenerationException($"資料來源回應內容為空。URL：{url}");
            }

            return body;
        }
    }

    private static string ApplyPlaceholders(string url, IDictionary<string, string> parameters)
    {
        var result = url;
        foreach (var kv in parameters)
        {
            result = result.Replace("{" + kv.Key + "}", WebUtility.UrlEncode(kv.Value));
        }
        return result;
    }

    private static string Truncate(string text, int maxLength)
    {
        return text.Length <= maxLength ? text : text[..maxLength] + "...";
    }
}
