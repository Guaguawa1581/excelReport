using System.IO.Compression;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using excelReport.Models;

namespace excelReport.Services;

/// <summary>用 OpenXML 讀取範本指定工作表的儲存格文字，正則掃出 {{xxx}} 標記並分類。</summary>
public class ReportTemplateScanner : IReportTemplateScanner
{
    /// <summary>
    /// 固定欄位標記語法 {{fieldName}}，fieldName 允許英數字、底線與點號（集合欄位）。
    /// </summary>
    private static readonly Regex MarkerRegex = new(@"\{\{\s*([a-zA-Z0-9_.]+)\s*\}\}", RegexOptions.Compiled);

    /// <summary>圖片標記語法 [[fieldName]]，與一般文字標記的 {{}} 語法分開，
    /// MiniExcel 完全不認得這個語法、套版時會原樣留著，等 ReportEngine 另外處理。</summary>
    private static readonly Regex ImageMarkerRegex = new(@"\[\[\s*([a-zA-Z0-9_.]+)\s*\]\]", RegexOptions.Compiled);

    /// <summary>保留欄位名，引擎會自動補上空字串值，不需要（也不應該）讓使用者手動映射。</summary>
    private const string ReservedColumnName = "_str";

    private readonly IConfiguration _configuration;

    public ReportTemplateScanner(IConfiguration configuration)
    {
        _configuration = configuration;
    }

    public TemplateScanResult Scan(byte[] templateBytes, string? sheet = null)
    {
        var fields = new List<string>();
        var collections = new Dictionary<string, List<string>>();
        var imageFields = new List<string>();

        using var stream = new MemoryStream(templateBytes);
        using var document = SpreadsheetDocument.Open(stream, false);

        var workbookPart = document.WorkbookPart
            ?? throw new ReportGenerationException("範本檔案格式不正確，找不到活頁簿內容。");

        var sharedStrings = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
        // 取特定sheet
        var worksheetPart = GetTargetWorksheetPart(workbookPart, sheet);

        var worksheetParts = workbookPart.WorksheetParts;
        // foreach (var worksheetPart in worksheetParts)
        // {

        var cells = worksheetPart.Worksheet.Descendants<Cell>();
        foreach (var cell in cells)
        {
            var text = GetCellText(cell, sharedStrings);
            if (string.IsNullOrEmpty(text)) continue;
            // 找出有"{{ }}"佔位符
            foreach (Match match in MarkerRegex.Matches(text))
            {
                var marker = match.Groups[1].Value;
                var dotIndex = marker.IndexOf('.');
                if (dotIndex < 0)
                {
                    if (marker == ReservedColumnName) continue;
                    if (!fields.Contains(marker)) fields.Add(marker);
                }
                else
                {
                    var collectionName = marker[..dotIndex];
                    var columnName = marker[(dotIndex + 1)..];
                    if (columnName == ReservedColumnName) continue;
                    if (!collections.TryGetValue(collectionName, out var columns))
                    {
                        columns = new List<string>();
                        collections[collectionName] = columns;
                    }
                    if (!columns.Contains(columnName)) columns.Add(columnName);
                }
            }
            // 找出有"[[ ]]"佔位符
            foreach (Match match in ImageMarkerRegex.Matches(text))
            {
                var marker = match.Groups[1].Value;
                if (!imageFields.Contains(marker)) imageFields.Add(marker);
            }
        }
        // }
        return new TemplateScanResult
        {
            Fields = fields,
            ImageFields = imageFields,
            Collections = collections
                .Select(kv => new CollectionScanResult { Name = kv.Key, Columns = kv.Value })
                .ToList()
        };
    }

    /// <summary>範本是否為「嚴格開放的 XML 試算表」(Strict Open XML Spreadsheet) 格式。
    /// MiniExcel 套版引擎不支援這個格式，即使掃描能成功，套版時也一定會丟出 NullReferenceException。
    /// 判斷依據：xl/workbook.xml 根節點的 conformance="strict" 屬性（ECMA-376 定義的格式旗標，
    /// 一般 Excel 活頁簿不會有這個屬性）。</summary>
    public bool IsStrictOpenXml(byte[] templateBytes)
    {
        using var stream = new MemoryStream(templateBytes);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
        var workbookEntry = archive.GetEntry("xl/workbook.xml");
        if (workbookEntry == null) return false;

        using var entryStream = workbookEntry.Open();
        using var reader = new StreamReader(entryStream, System.Text.Encoding.UTF8);
        var content = reader.ReadToEnd();
        return content.Contains("conformance=\"strict\"", StringComparison.Ordinal);
    }

    /// <summary>決定要處理哪個工作表：數字視為 0 起始的順序索引，字串視為工作表名稱。
    /// sheet 參數是每份範本（主範本或子報表）各自存的選擇；沒填的話 fallback 到 appsettings 的
    /// TemplateScan:Sheet（舊設定檔沒有這個欄位時的相容行為），再沒有就預設第一個（索引 0）。</summary>
    public WorksheetPart GetTargetWorksheetPart(WorkbookPart workbookPart, string? sheet = null)
    {
        var sheets = workbookPart.Workbook.Sheets?.Elements<Sheet>().ToList();
        if (sheets == null || sheets.Count == 0)
        {
            throw new ReportGenerationException("範本檔案格式不正確，找不到工作表。");
        }

        var setting = !string.IsNullOrWhiteSpace(sheet) ? sheet : (_configuration["TemplateScan:Sheet"] ?? "0");

        Sheet targetSheet;
        if (int.TryParse(setting, out var index))
        {
            if (index < 0 || index >= sheets.Count)
            {
                index = 0; // 超過就用0
            }
            targetSheet = sheets[index];
        }
        else
        {
            targetSheet = sheets.FirstOrDefault(s => s.Name == setting)
                ?? sheets[index];
        }

        var sheetId = targetSheet.Id?.Value
            ?? throw new ReportGenerationException("範本檔案格式不正確，工作表缺少 Id。");

        return (WorksheetPart)workbookPart.GetPartById(sheetId);
    }

    /// <summary>依原始 xlsx 內的排列順序回傳所有工作表名稱，供設定頁面畫出「(索引)名稱」的
    /// 選單，讓使用者自己選要用哪一頁，取代原本寫死在 appsettings 的單一預設值。</summary>
    public List<string> GetSheetNames(byte[] templateBytes)
    {
        using var stream = new MemoryStream(templateBytes);
        using var document = SpreadsheetDocument.Open(stream, false);
        var workbookPart = document.WorkbookPart
            ?? throw new ReportGenerationException("範本檔案格式不正確，找不到活頁簿內容。");

        return workbookPart.Workbook.Sheets?.Elements<Sheet>()
            .Select(s => s.Name?.Value ?? "")
            .ToList() ?? new List<string>();
    }

    private static string GetCellText(Cell cell, SharedStringTablePart? sharedStrings)
    {
        if (cell.DataType?.Value == CellValues.SharedString)
        {
            if (sharedStrings == null || cell.CellValue == null) return "";
            if (int.TryParse(cell.CellValue.Text, out var idx))
            {
                var items = sharedStrings.SharedStringTable.Elements<SharedStringItem>().ToList();
                if (idx >= 0 && idx < items.Count)
                {
                    return items[idx].InnerText;
                }
            }
            return "";
        }

        if (cell.DataType?.Value == CellValues.InlineString)
        {
            return cell.InlineString?.InnerText ?? "";
        }

        return cell.CellValue?.Text ?? "";
    }
}
