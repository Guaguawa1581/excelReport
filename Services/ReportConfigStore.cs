using excelReport.Models;
using Newtonsoft.Json;
using Newtonsoft.Json.Serialization;

namespace excelReport.Services;

/// <summary>以磁碟 JSON 檔實作的報表設定存取；不使用資料庫。</summary>
public class ReportConfigStore : IReportConfigStore
{
    private readonly string _configsDir;
    private readonly ILogger<ReportConfigStore> _logger;
    private static readonly JsonSerializerSettings JsonSettings = new()
    {
        ContractResolver = new CamelCasePropertyNamesContractResolver(),
        Formatting = Formatting.Indented
    };
    private static readonly object FileLock = new();

    public ReportConfigStore(IWebHostEnvironment env, ILogger<ReportConfigStore> logger)
    {
        _configsDir = Path.Combine(env.ContentRootPath, "App_Data", "configs");
        _logger = logger;
        Directory.CreateDirectory(_configsDir);
        MigrateLegacyFiles();
    }

    /// <summary>舊版設定檔以 {code}.json 命名、沒有 LogId。啟動時掃描一次，補上 LogId 並把檔案
    /// 搬遷成 {logId}.json，之後 Code 改名就只是改內容、不用搬動實體檔案。</summary>
    private void MigrateLegacyFiles()
    {
        lock (FileLock)
        {
            foreach (var file in Directory.EnumerateFiles(_configsDir, "*.json"))
            {
                ReportConfig? config;
                try
                {
                    config = ReadFile(file);
                }
                catch (JsonException ex)
                {
                    _logger.LogWarning(ex, "設定檔 {File} 格式錯誤，略過遷移", file);
                    continue;
                }
                if (config == null || !string.IsNullOrWhiteSpace(config.LogId)) continue;

                config.LogId = Guid.NewGuid().ToString();
                var newPath = PathFor(config.LogId);
                File.WriteAllText(newPath, JsonConvert.SerializeObject(config, JsonSettings));
                if (!string.Equals(newPath, file, StringComparison.OrdinalIgnoreCase))
                {
                    File.Delete(file);
                }
                _logger.LogInformation("設定檔 {File} 補上 logId 並搬遷為 {NewFile}", file, newPath);
            }
        }
    }

    public List<ReportConfig> GetAll()
    {
        lock (FileLock)
        {
            var result = new List<ReportConfig>();
            foreach (var file in Directory.EnumerateFiles(_configsDir, "*.json").OrderBy(f => f))
            {
                // 單一設定檔格式錯誤（例如手動編輯壞掉）只跳過該檔案並記警告，
                // 不能讓整個網站因為一份壞掉的設定檔而無法啟動。
                ReportConfig? config;
                try
                {
                    config = ReadFile(file);
                }
                catch (JsonException ex)
                {
                    _logger.LogWarning(ex, "設定檔 {File} 格式錯誤，已略過", file);
                    continue;
                }
                if (config != null) result.Add(config);
            }
            return result;
        }
    }

    public ReportConfig? Get(string code)
    {
        lock (FileLock)
        {
            return GetAll().FirstOrDefault(c => c.Code == code);
        }
    }

    public void Save(ReportConfig config)
    {
        lock (FileLock)
        {
            if (string.IsNullOrWhiteSpace(config.LogId))
            {
                config.LogId = Guid.NewGuid().ToString();
            }
            var path = PathFor(config.LogId);
            var json = JsonConvert.SerializeObject(config, JsonSettings);
            File.WriteAllText(path, json);
        }
    }

    public void Delete(string code)
    {
        lock (FileLock)
        {
            var config = GetAll().FirstOrDefault(c => c.Code == code);
            if (config == null) return;
            var path = PathFor(config.LogId);
            if (File.Exists(path)) File.Delete(path);
        }
    }

    public bool Exists(string code)
    {
        lock (FileLock)
        {
            return GetAll().Any(c => c.Code == code);
        }
    }

    private string PathFor(string logId) => Path.Combine(_configsDir, $"{logId}.json");

    private static ReportConfig? ReadFile(string path)
    {
        var json = File.ReadAllText(path);
        return JsonConvert.DeserializeObject<ReportConfig>(json, JsonSettings);
    }
}
