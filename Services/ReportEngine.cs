using System.Diagnostics;
using excelReport.Models;
using Microsoft.Extensions.Caching.Memory;
using MiniExcelLibs;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace excelReport.Services;

public class ReportEngine : IReportEngine
{
    private readonly IReportConfigStore _configStore;
    private readonly IDataSourceClient _dataSourceClient;
    private readonly IWebHostEnvironment _env;
    private readonly IMemoryCache _cache;
    private readonly ILogger<ReportEngine> _logger;

    public ReportEngine(
        IReportConfigStore configStore,
        IDataSourceClient dataSourceClient,
        IWebHostEnvironment env,
        IMemoryCache cache,
        ILogger<ReportEngine> logger)
    {
        _configStore = configStore;
        _dataSourceClient = dataSourceClient;
        _env = env;
        _cache = cache;
        _logger = logger;
    }

    public async Task<ReportGenerationResult> GenerateAsync(string code, IDictionary<string, string> parameters, CancellationToken cancellationToken = default)
    {
        var stopwatch = Stopwatch.StartNew();
        ReportConfig? config = null;
        try
        {
            config = _configStore.Get(code)
                ?? throw new ReportGenerationException($"找不到報表設定：{code}");

            ValidateParameters(config, parameters);

            var json = await _dataSourceClient.FetchAsync(config.DataSource, parameters, cancellationToken);

            JObject root;
            try
            {
                root = JObject.Parse(json);
            }
            catch (JsonException ex)
            {
                throw new ReportGenerationException($"資料來源回傳的內容不是有效的 JSON：{ex.Message}", ex);
            }

            var rootToken = ResolveRoot(root, config.Mapping.Root);
            var data = BuildDataDictionary(config.Mapping, rootToken);

            var fileBytes = await RenderTemplateAsync(config, data, cancellationToken);
            var fileName = BuildFileName(config, parameters);

            stopwatch.Stop();
            WriteLog(config.Code, parameters, stopwatch.ElapsedMilliseconds, true, null);

            return new ReportGenerationResult { FileBytes = fileBytes, FileName = fileName };
        }
        catch (ReportGenerationException ex)
        {
            stopwatch.Stop();
            WriteLog(config?.Code ?? code, parameters, stopwatch.ElapsedMilliseconds, false, ex.Message);
            throw;
        }
        catch (Exception ex)
        {
            stopwatch.Stop();
            _logger.LogError(ex, "產生報表 {Code} 時發生未預期錯誤", code);
            WriteLog(config?.Code ?? code, parameters, stopwatch.ElapsedMilliseconds, false, ex.Message);
            throw new ReportGenerationException($"產生報表時發生未預期的錯誤：{ex.Message}", ex);
        }
    }

    private static void ValidateParameters(ReportConfig config, IDictionary<string, string> parameters)
    {
        var missing = config.Parameters
            .Where(p => p.Required && (!parameters.TryGetValue(p.Key, out var v) || string.IsNullOrWhiteSpace(v)))
            .Select(p => p.Label)
            .ToList();

        if (missing.Count > 0)
        {
            throw new ReportGenerationException($"缺少必填參數：{string.Join("、", missing)}");
        }
    }

    private static JToken ResolveRoot(JObject root, string rootPath)
    {
        if (string.IsNullOrWhiteSpace(rootPath) || rootPath == "$")
        {
            return root;
        }

        var token = root.SelectToken(rootPath);
        if (token == null)
        {
            throw new ReportGenerationException($"找不到資料根節點，JSONPath：{rootPath}");
        }
        return token;
    }

    /// <summary>
    /// 保留欄位名，永遠對應空字串。範本中每個單值標記後面都會附加 {{_str}}（或集合內的
    /// {{collection._str}}），讓儲存格內含多個標記，藉此讓 MiniExcel 套版時保留文字型別，
    /// 避免像 "0040" 這種數字外觀的字串被誤判為數值而遺失前導零。
    /// </summary>
    private const string ReservedColumnName = "_str";

    private static Dictionary<string, object> BuildDataDictionary(MappingConfig mapping, JToken rootToken)
    {
        var data = new Dictionary<string, object> { [ReservedColumnName] = "" };

        foreach (var (fieldKey, path) in mapping.Fields)
        {
            var token = rootToken.SelectToken(path);
            if (token == null)
            {
                throw new ReportGenerationException($"欄位「{fieldKey}」的 JSONPath 找不到對應節點：{path}");
            }
            data[fieldKey] = TokenToValue(token);
        }

        foreach (var collection in mapping.Collections)
        {
            var arrayToken = rootToken.SelectToken(collection.Path);
            if (arrayToken == null)
            {
                throw new ReportGenerationException($"集合「{collection.Name}」的 JSONPath 找不到對應節點：{collection.Path}");
            }
            if (arrayToken.Type != JTokenType.Array)
            {
                throw new ReportGenerationException($"集合「{collection.Name}」對應的節點不是陣列：{collection.Path}");
            }

            var rows = new List<IDictionary<string, object>>();
            foreach (var item in arrayToken)
            {
                var row = new Dictionary<string, object> { [ReservedColumnName] = "" };
                foreach (var (columnKey, columnPath) in collection.Columns)
                {
                    var valueToken = item.SelectToken(columnPath);
                    if (valueToken == null)
                    {
                        throw new ReportGenerationException(
                            $"集合「{collection.Name}」欄位「{columnKey}」的 JSONPath 找不到對應節點：{columnPath}");
                    }
                    row[columnKey] = TokenToValue(valueToken);
                }

                if (collection.Spread != null)
                {
                    ApplySpread(item, collection.Spread, row);
                }

                rows.Add(row);
            }

            data[collection.Name] = rows;
        }

        return data;
    }

    private static void ApplySpread(JToken item, SpreadConfig spread, Dictionary<string, object> row)
    {
        var spreadToken = item.SelectToken(spread.From);
        var values = spreadToken != null && spreadToken.Type == JTokenType.Array
            ? spreadToken.Select(TokenToValue).ToList()
            : new List<object>();

        for (var i = 0; i < spread.Max; i++)
        {
            var key = $"{spread.Prefix}{i + 1}";
            row[key] = i < values.Count ? values[i] : "";
        }
    }

    private static object TokenToValue(JToken token)
    {
        if (token.Type == JTokenType.Null) return "";
        if (token is JValue jv) return jv.Value?.ToString() ?? "";
        return token.ToString();
    }

    private async Task<byte[]> RenderTemplateAsync(ReportConfig config, Dictionary<string, object> data, CancellationToken cancellationToken)
    {
        var templatePath = Path.Combine(_env.ContentRootPath, "App_Data", "templates", config.TemplateFile);
        if (!File.Exists(templatePath))
        {
            throw new ReportGenerationException($"找不到範本檔案：{config.TemplateFile}");
        }

        var lastWriteTicks = File.GetLastWriteTimeUtc(templatePath).Ticks;
        var cacheKey = $"tpl:{config.TemplateFile}:{lastWriteTicks}";
        var templateBytes = await _cache.GetOrCreateAsync(cacheKey, entry =>
        {
            entry.SlidingExpiration = TimeSpan.FromMinutes(30);
            return Task.FromResult(File.ReadAllBytes(templatePath));
        }) ?? throw new ReportGenerationException("範本檔案讀取失敗。");

        using var outputStream = new MemoryStream();
        try
        {
            await outputStream.SaveAsByTemplateAsync(templateBytes, data, cancellationToken: cancellationToken);
        }
        catch (Exception ex) when (ex is not ReportGenerationException)
        {
            throw new ReportGenerationException($"套版產生 xlsx 時發生錯誤：{ex.Message}", ex);
        }

        return outputStream.ToArray();
    }

    private static string BuildFileName(ReportConfig config, IDictionary<string, string> parameters)
    {
        var keyParam = parameters.TryGetValue("docNo", out var docNo) && !string.IsNullOrWhiteSpace(docNo)
            ? docNo
            : parameters.Values.FirstOrDefault(v => !string.IsNullOrWhiteSpace(v)) ?? "";

        var namePart = string.IsNullOrWhiteSpace(keyParam) ? config.Name : $"{config.Name}_{keyParam}";
        return $"{namePart}_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
    }

    private void WriteLog(string code, IDictionary<string, string> parameters, long elapsedMs, bool success, string? error)
    {
        try
        {
            var logsDir = Path.Combine(_env.ContentRootPath, "App_Data", "logs");
            Directory.CreateDirectory(logsDir);
            var logFile = Path.Combine(logsDir, $"report-{DateTime.Now:yyyyMMdd}.log");

            var entry = new
            {
                time = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
                code,
                parameters,
                elapsedMs,
                success,
                error
            };

            File.AppendAllText(logFile, JsonConvert.SerializeObject(entry) + Environment.NewLine);
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "寫入報表產生記錄檔失敗");
        }
    }
}
