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
            var path = PathFor(code);
            if (!File.Exists(path)) return null;
            try
            {
                return ReadFile(path);
            }
            catch (JsonException ex)
            {
                _logger.LogWarning(ex, "設定檔 {File} 格式錯誤", path);
                return null;
            }
        }
    }

    public void Save(ReportConfig config)
    {
        lock (FileLock)
        {
            var path = PathFor(config.Code);
            var json = JsonConvert.SerializeObject(config, JsonSettings);
            File.WriteAllText(path, json);
        }
    }

    public void Delete(string code)
    {
        lock (FileLock)
        {
            var path = PathFor(code);
            if (File.Exists(path)) File.Delete(path);
        }
    }

    public bool Exists(string code)
    {
        lock (FileLock)
        {
            return File.Exists(PathFor(code));
        }
    }

    private string PathFor(string code) => Path.Combine(_configsDir, $"{code}.json");

    private static ReportConfig? ReadFile(string path)
    {
        var json = File.ReadAllText(path);
        return JsonConvert.DeserializeObject<ReportConfig>(json, JsonSettings);
    }
}
