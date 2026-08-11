using excelReport.Models;
using excelReport.Services;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Net.Http.Headers;

namespace excelReport.Controllers;

/// <summary>報表產生頁面：選擇報表 → 輸入參數 → 下載 xlsx。</summary>
public class ReportController : Controller
{
    private readonly IReportConfigStore _configStore;
    private readonly IReportEngine _reportEngine;
    private readonly ILogger<ReportController> _logger;

    public ReportController(IReportConfigStore configStore, IReportEngine reportEngine, ILogger<ReportController> logger)
    {
        _configStore = configStore;
        _reportEngine = reportEngine;
        _logger = logger;
    }

    [HttpGet]
    public IActionResult Index()
    {
        var items = _configStore.GetAll()
            .Select(c => new ReportListItemViewModel { Code = c.Code, Name = c.Name, Parameters = c.Parameters })
            .ToList();
        return View(items);
    }

    [HttpGet]
    public async Task<IActionResult> Generate(string code, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(code))
        {
            return RenderError("請選擇要產生的報表。");
        }

        var parameters = new Dictionary<string, string>();
        foreach (var key in Request.Query.Keys)
        {
            if (key.Equals("code", StringComparison.OrdinalIgnoreCase)) continue;
            parameters[key] = Request.Query[key].ToString();
        }

        try
        {
            var result = await _reportEngine.GenerateAsync(code, parameters, cancellationToken);
            var contentDisposition = new ContentDispositionHeaderValue("attachment");
            contentDisposition.SetHttpFileName(result.FileName);
            Response.Headers.Append("Content-Disposition", contentDisposition.ToString());
            return File(result.FileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        }
        catch (ReportGenerationException ex)
        {
            return RenderError(ex.Message);
        }
    }

    private IActionResult RenderError(string message)
    {
        Response.StatusCode = 400;
        return View("Error", new ReportErrorViewModel { Title = "報表產生失敗", Message = message });
    }
}
