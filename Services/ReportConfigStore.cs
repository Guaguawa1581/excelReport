using excelReport.Models;
using Newtonsoft.Json;
using Newtonsoft.Json.Serialization;

namespace excelReport.Services;

/// <summary>以磁碟 JSON 檔實作的報表設定存取；不使用資料庫。</summary>
public class ReportConfigStore : IReportConfigStore
{
    private readonly string _configsDir;
    private static readonly JsonSerializerSettings JsonSettings = new()
    {
        ContractResolver = new CamelCasePropertyNamesContractResolver(),
        Formatting = Formatting.Indented
    };
    private static readonly object FileLock = new();

    public ReportConfigStore(IWebHostEnvironment env)
    {
        _configsDir = Path.Combine(env.ContentRootPath, "App_Data", "configs");
        Directory.CreateDirectory(_configsDir);
    }

    public List<ReportConfig> GetAll()
    {
        lock (FileLock)
        {
            var result = new List<ReportConfig>();
            foreach (var file in Directory.EnumerateFiles(_configsDir, "*.json").OrderBy(f => f))
            {
                var config = ReadFile(file);
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
            return File.Exists(path) ? ReadFile(path) : null;
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
