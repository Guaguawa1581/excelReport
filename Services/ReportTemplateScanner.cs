using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using excelReport.Models;

namespace excelReport.Services;

/// <summary>用 OpenXML 讀取範本所有工作表的儲存格文字，正則掃出 {{xxx}} 標記並分類。</summary>
public class ReportTemplateScanner : IReportTemplateScanner
{
    private static readonly Regex MarkerRegex = new(@"\{\{\s*([a-zA-Z0-9_.]+)\s*\}\}", RegexOptions.Compiled);

    /// <summary>保留欄位名，引擎會自動補上空字串值，不需要（也不應該）讓使用者手動映射。</summary>
    private const string ReservedColumnName = "_str";

    public TemplateScanResult Scan(byte[] templateBytes)
    {
        var fields = new List<string>();
        var collections = new Dictionary<string, List<string>>();

        using var stream = new MemoryStream(templateBytes);
        using var document = SpreadsheetDocument.Open(stream, false);

        var workbookPart = document.WorkbookPart
            ?? throw new ReportGenerationException("範本檔案格式不正確，找不到活頁簿內容。");

        var sharedStrings = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();

        foreach (var worksheetPart in workbookPart.WorksheetParts)
        {
            var cells = worksheetPart.Worksheet.Descendants<Cell>();
            foreach (var cell in cells)
            {
                var text = GetCellText(cell, sharedStrings);
                if (string.IsNullOrEmpty(text)) continue;

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
            }
        }

        return new TemplateScanResult
        {
            Fields = fields,
            Collections = collections
                .Select(kv => new CollectionScanResult { Name = kv.Key, Columns = kv.Value })
                .ToList()
        };
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
