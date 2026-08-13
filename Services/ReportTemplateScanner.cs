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

    public TemplateScanResult Scan(byte[] templateBytes)
    {
        var fields = new List<string>();
        var collections = new Dictionary<string, List<string>>();
        var imageFields = new List<string>();

        using var stream = new MemoryStream(templateBytes);
        using var document = SpreadsheetDocument.Open(stream, false);

        var workbookPart = document.WorkbookPart
            ?? throw new ReportGenerationException("範本檔案格式不正確，找不到活頁簿內容。");

        var sharedStrings = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
        var worksheetPart = GetTargetWorksheetPart(workbookPart);

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

        return new TemplateScanResult
        {
            Fields = fields,
            ImageFields = imageFields,
            Collections = collections
                .Select(kv => new CollectionScanResult { Name = kv.Key, Columns = kv.Value })
                .ToList()
        };
    }

    /// <summary>依 appsettings 的 TemplateScan:Sheet 決定要掃描哪個工作表：
    /// 數字視為 0 起始的順序索引，字串視為工作表名稱。未設定時預設第一個（索引 0）。</summary>
    private WorksheetPart GetTargetWorksheetPart(WorkbookPart workbookPart)
    {
        var sheets = workbookPart.Workbook.Sheets?.Elements<Sheet>().ToList();
        if (sheets == null || sheets.Count == 0)
        {
            throw new ReportGenerationException("範本檔案格式不正確，找不到工作表。");
        }

        var setting = _configuration["TemplateScan:Sheet"] ?? "0";

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
