using System.Diagnostics;
using excelReport.Models;
using Microsoft.AspNetCore.Mvc;

namespace excelReport.Controllers;

/// <summary>僅保留全域例外處理的備援錯誤頁面（非預期例外時使用）。</summary>
public class HomeController : Controller
{
    [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
    public IActionResult Error()
    {
        return View("Error", new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
    }
}
