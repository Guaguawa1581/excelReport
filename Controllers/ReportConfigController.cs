using excelReport.Models;
using excelReport.Services;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;

namespace excelReport.Controllers;

/// <summary>報表設定管理頁面：清單、新增/編輯/刪除/複製、範本上傳掃描、測試連線。</summary>
public class ReportConfigController : Controller
{
    private readonly IReportConfigStore _configStore;
    private readonly IReportTemplateScanner _scanner;
    private readonly IDataSourceClient _dataSourceClient;
    private readonly IWebHostEnvironment _env;
    private readonly ILogger<ReportConfigController> _logger;

    public ReportConfigController(
        IReportConfigStore configStore,
        IReportTemplateScanner scanner,
        IDataSourceClient dataSourceClient,
        IWebHostEnvironment env,
        ILogger<ReportConfigController> logger)
    {
        _configStore = configStore;
        _scanner = scanner;
        _dataSourceClient = dataSourceClient;
        _env = env;
        _logger = logger;
    }

    private string TemplatesDir => Path.Combine(_env.ContentRootPath, "App_Data", "templates");

    [HttpGet]
    public IActionResult Index()
    {
        var configs = _configStore.GetAll();
        return View(configs);
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public IActionResult Delete(string code)
    {
        _configStore.Delete(code);
        TempData["SuccessMessage"] = $"已刪除設定：{code}";
        return RedirectToAction(nameof(Index));
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public IActionResult Copy(string code)
    {
        var source = _configStore.Get(code);
        if (source == null)
        {
            TempData["ErrorMessage"] = $"找不到設定：{code}";
            return RedirectToAction(nameof(Index));
        }

        var newCode = $"{source.Code}_COPY";
        var suffix = 1;
        while (_configStore.Exists(newCode))
        {
            newCode = $"{source.Code}_COPY{++suffix}";
        }

        source.Code = newCode;
        source.Name = $"{source.Name} (複製)";
        _configStore.Save(source);

        TempData["SuccessMessage"] = $"已複製為：{newCode}";
        return RedirectToAction(nameof(Edit), new { code = newCode });
    }

    [HttpGet]
    public IActionResult Edit(string? code, string? template)
    {
        var vm = new ReportConfigEditViewModel
        {
            AvailableTemplates = GetAvailableTemplates()
        };

        if (string.IsNullOrWhiteSpace(code) || code == "new")
        {
            vm.IsNew = true;
            vm.Config = new ReportConfig
            {
                DataSource = new DataSourceConfig { Method = "GET", TimeoutSec = 30 }
            };
        }
        else
        {
            var existing = _configStore.Get(code);
            if (existing == null)
            {
                TempData["ErrorMessage"] = $"找不到設定：{code}";
                return RedirectToAction(nameof(Index));
            }
            vm.IsNew = false;
            vm.OriginalCode = code;
            vm.Config = existing;
        }

        // 允許透過 template 參數切換要掃描/預覽的範本，而不影響尚未儲存的設定。
        if (!string.IsNullOrWhiteSpace(template))
        {
            vm.Config.TemplateFile = template;
        }

        if (!string.IsNullOrWhiteSpace(vm.Config.TemplateFile))
        {
            vm.ScanResult = TryScanTemplate(vm.Config.TemplateFile);
        }

        return View(vm);
    }

    [HttpPost]
    [ActionName("Edit")]
    [ValidateAntiForgeryToken]
    public IActionResult EditPost(string configJson, string? originalCode)
    {
        ReportConfig? config;
        try
        {
            config = JsonConvert.DeserializeObject<ReportConfig>(configJson);
        }
        catch (JsonException ex)
        {
            return EditWithError(configJson, originalCode, $"送出的設定內容格式錯誤：{ex.Message}");
        }

        if (config == null || string.IsNullOrWhiteSpace(config.Code))
        {
            return EditWithError(configJson, originalCode, "設定代碼（code）不可為空。");
        }

        if (string.IsNullOrWhiteSpace(config.TemplateFile))
        {
            return EditWithError(configJson, originalCode, "請先上傳或選擇一份範本 xlsx。");
        }

        var scanResult = TryScanTemplate(config.TemplateFile);
        if (scanResult == null)
        {
            return EditWithError(configJson, originalCode, $"找不到範本檔案：{config.TemplateFile}");
        }

        var isRename = !string.IsNullOrWhiteSpace(originalCode) && originalCode != config.Code;
        if (isRename && _configStore.Exists(config.Code))
        {
            return EditWithError(configJson, originalCode, $"設定代碼「{config.Code}」已存在，請換一個代碼。");
        }

        _configStore.Save(config);
        if (isRename)
        {
            _configStore.Delete(originalCode!);
        }

        TempData["SuccessMessage"] = "設定已儲存。";
        return RedirectToAction(nameof(Edit), new { code = config.Code });
    }

    [HttpGet]
    public IActionResult DownloadTemplate(string fileName)
    {
        if (string.IsNullOrWhiteSpace(fileName))
        {
            return NotFound();
        }

        var safeFileName = Path.GetFileName(fileName);
        var path = Path.Combine(TemplatesDir, safeFileName);
        if (!System.IO.File.Exists(path))
        {
            return NotFound();
        }

        var bytes = System.IO.File.ReadAllBytes(path);
        return File(bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", safeFileName);
    }

    [HttpPost]
    public async Task<IActionResult> UploadTemplate(IFormFile file)
    {
        if (file == null || file.Length == 0)
        {
            return BadRequest(new { message = "請選擇要上傳的檔案。" });
        }

        if (!Path.GetExtension(file.FileName).Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
        {
            return BadRequest(new { message = "只接受 .xlsx 檔案。" });
        }

        Directory.CreateDirectory(TemplatesDir);
        var safeFileName = SanitizeFileName(file.FileName);
        var path = Path.Combine(TemplatesDir, safeFileName);

        byte[] bytes;
        using (var ms = new MemoryStream())
        {
            await file.CopyToAsync(ms);
            bytes = ms.ToArray();
        }

        if (_scanner.IsStrictOpenXml(bytes))
        {
            return BadRequest(new
            {
                message = "此範本是用「嚴格開放的 XML 試算表 (Strict Open XML Spreadsheet)」格式儲存，套版引擎不支援此格式。" +
                    "請在 Excel 開啟後用「另存新檔」，檔案類型選擇一般的「Excel 活頁簿 (*.xlsx)」（不是 Strict Open XML Spreadsheet）再重新上傳。"
            });
        }

        TemplateScanResult scanResult;
        try
        {
            scanResult = _scanner.Scan(bytes);
        }
        catch (ReportGenerationException ex)
        {
            return BadRequest(new { message = ex.Message });
        }

        await System.IO.File.WriteAllBytesAsync(path, bytes);

        return Ok(new UploadTemplateResult { FileName = safeFileName, ScanResult = scanResult });
    }

    [HttpPost]
    public async Task<IActionResult> TestConnection([FromBody] TestConnectionRequest request, CancellationToken cancellationToken)
    {
        var dataSource = new DataSourceConfig
        {
            Type = request.Type,
            Url = request.Url,
            Method = request.Method,
            Headers = request.Headers,
            TimeoutSec = request.TimeoutSec > 0 ? request.TimeoutSec : 30,
            StaticJson = request.StaticJson
        };

        try
        {
            var body = await _dataSourceClient.FetchAsync(dataSource, request.Parameters, cancellationToken);
            return Ok(new TestConnectionResponse { Success = true, Body = body });
        }
        catch (ReportGenerationException ex)
        {
            return Ok(new TestConnectionResponse { Success = false, ErrorMessage = ex.Message });
        }
    }

    private IActionResult EditWithError(string configJson, string? originalCode, string errorMessage)
    {
        ReportConfig config;
        try
        {
            config = JsonConvert.DeserializeObject<ReportConfig>(configJson) ?? new ReportConfig();
        }
        catch
        {
            config = new ReportConfig();
        }

        var vm = new ReportConfigEditViewModel
        {
            IsNew = string.IsNullOrWhiteSpace(originalCode),
            OriginalCode = originalCode ?? "",
            Config = config,
            ScanResult = string.IsNullOrWhiteSpace(config.TemplateFile) ? null : TryScanTemplate(config.TemplateFile),
            AvailableTemplates = GetAvailableTemplates(),
            ErrorMessage = errorMessage
        };
        return View(vm);
    }

    private TemplateScanResult? TryScanTemplate(string fileName)
    {
        if (string.IsNullOrWhiteSpace(fileName)) return null;
        var path = Path.Combine(TemplatesDir, fileName);
        if (!System.IO.File.Exists(path)) return null;

        try
        {
            return _scanner.Scan(System.IO.File.ReadAllBytes(path));
        }
        catch (ReportGenerationException ex)
        {
            _logger.LogWarning("掃描範本 {File} 失敗：{Message}", fileName, ex.Message);
            return null;
        }
    }

    private List<string> GetAvailableTemplates()
    {
        if (!Directory.Exists(TemplatesDir)) return new List<string>();
        return Directory.EnumerateFiles(TemplatesDir, "*.xlsx")
            .Select(Path.GetFileName)
            .Where(f => f != null)
            .Select(f => f!)
            .OrderBy(f => f)
            .ToList();
    }

    private static string SanitizeFileName(string fileName)
    {
        var name = Path.GetFileName(fileName);
        foreach (var c in Path.GetInvalidFileNameChars())
        {
            name = name.Replace(c, '_');
        }
        return name;
    }
}
