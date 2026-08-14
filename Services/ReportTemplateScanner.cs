using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
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
    public WorksheetPart GetTargetWorksheetPart(WorkbookPart workbookPart)
    {
        var sheets = workbookPart.Workbook.Sheets?.Elements<Sheet>().ToList();
        if (sheets == null || sheets.Count == 0)
        {
            throw new ReportGenerationException("範本檔案格式不正確，找不到工作表。");
        }

        var targetSheet = ResolveTargetSheet(sheets);
        var sheetId = targetSheet.Id?.Value
            ?? throw new ReportGenerationException("範本檔案格式不正確，工作表缺少 Id。");

        return (WorksheetPart)workbookPart.GetPartById(sheetId);
    }

    private Sheet ResolveTargetSheet(List<Sheet> sheets)
    {
        var setting = _configuration["TemplateScan:Sheet"] ?? "0";

        if (int.TryParse(setting, out var index))
        {
            if (index < 0 || index >= sheets.Count)
            {
                index = 0; // 超過就用0
            }
            return sheets[index];
        }

        return sheets.FirstOrDefault(s => s.Name == setting) ?? sheets[0];
    }

    /// <summary>
    /// MiniExcel 套版時是對整個活頁簿下手，不會只處理 Scan() 讀取的那個工作表。範本裡若還有
    /// 其他分頁殘留 {{}} 標記（例如舊草稿、其他報表共用同一個檔案時留下的分頁），那些標記不會
    /// 出現在 Scan() 建立的資料字典裡，套版時 MiniExcel 找不到對應 key 就會丟出
    /// NullReferenceException，讓整份報表產生失敗。這裡在套版前先把目標工作表以外的分頁整個砍掉，
    /// 讓 MiniExcel 實際處理的內容範圍跟 Scan() 掃描的範圍一致。
    /// </summary>
    public byte[] IsolateTargetSheet(byte[] templateBytes)
    {
        HashSet<string> removedRelationshipIds;

        using var stream = new MemoryStream();
        stream.Write(templateBytes, 0, templateBytes.Length);
        stream.Position = 0;

        using (var document = SpreadsheetDocument.Open(stream, true))
        {
            var workbookPart = document.WorkbookPart
                ?? throw new ReportGenerationException("範本檔案格式不正確，找不到活頁簿內容。");

            var sheetsElement = workbookPart.Workbook.Sheets;
            var sheets = sheetsElement?.Elements<Sheet>().ToList();
            if (sheets == null || sheets.Count == 0)
            {
                throw new ReportGenerationException("範本檔案格式不正確，找不到工作表。");
            }

            if (sheets.Count == 1)
            {
                return templateBytes;
            }

            var targetSheet = ResolveTargetSheet(sheets);
            var originalIndex = (uint)sheets.IndexOf(targetSheet);
            var sheetsToRemove = sheets.Where(s => s != targetSheet).ToList();
            var removedNames = sheetsToRemove.Select(s => s.Name?.Value ?? "").Where(n => n.Length > 0).ToHashSet();
            removedRelationshipIds = sheetsToRemove.Select(s => s.Id?.Value ?? "").Where(id => id.Length > 0).ToHashSet();

            foreach (var sheet in sheetsToRemove)
            {
                var sheetId = sheet.Id?.Value;
                if (sheetId != null)
                {
                    var part = workbookPart.GetPartById(sheetId);
                    workbookPart.DeletePart(part);
                }
                sheet.Remove();
            }

            RemoveDanglingDefinedNames(workbookPart.Workbook.DefinedNames, originalIndex, removedNames);

            workbookPart.Workbook.Save();
        }

        // OpenXml SDK 3.0.x 的 DeletePart 已知不會清掉「被刪除分頁所屬工作簿」自己的關聯檔
        // （xl/_rels/workbook.xml.rels）：實體檔案、[Content_Types].xml 都會正確更新，唯獨這個
        // .rels 還留著指向已刪除檔案的 Relationship，導致這份 xlsx 用 OpenXml SDK 重新開啟時
        // 會拋出「Specified part does not exist in the package」（其他工具如 Excel 通常會容忍
        // 這種殘留參照，但我們後續 Scan()／嵌圖片都要用 SDK 重新開啟，所以必須手動補洗乾淨）。
        return RemoveDanglingWorkbookRelationships(stream.ToArray(), removedRelationshipIds);
    }

    /// <summary>把只含目標工作表的套版結果貼回完整的原始範本，取代目標工作表的內容，
    /// 其餘工作表（包含它們自己的樣式、圖片、關聯檔）完全不動。</summary>
    public byte[] MergeRenderedTargetSheet(byte[] originalTemplateBytes, byte[] renderedTargetSheetBytes)
    {
        Worksheet renderedWorksheet;
        using (var renderedStream = new MemoryStream(renderedTargetSheetBytes))
        using (var renderedDocument = SpreadsheetDocument.Open(renderedStream, false))
        {
            var renderedWorkbookPart = renderedDocument.WorkbookPart
                ?? throw new ReportGenerationException("套版結果格式不正確，找不到活頁簿內容。");
            var renderedSheet = renderedWorkbookPart.Workbook.Sheets?.Elements<Sheet>().FirstOrDefault()
                ?? throw new ReportGenerationException("套版結果格式不正確，找不到工作表。");
            var renderedWorksheetPart = (WorksheetPart)renderedWorkbookPart.GetPartById(renderedSheet.Id!.Value!);
            renderedWorksheet = (Worksheet)renderedWorksheetPart.Worksheet.CloneNode(true);
        }

        using var originalStream = new MemoryStream();
        originalStream.Write(originalTemplateBytes, 0, originalTemplateBytes.Length);
        originalStream.Position = 0;

        // 用 renderedTargetSheetBytes 只有一個 sheet 這件事來判斷「原本是不是只有一個 sheet」
        // 並不可靠（IsolateTargetSheet 對單一 sheet 範本本來就會直接原樣回傳，所以這種情況下
        // renderedTargetSheetBytes 其實已經是完整範本套版後的結果），所以這裡改成直接檢查原始
        // 範本本身的 sheet 數。
        using (var document = SpreadsheetDocument.Open(originalStream, true))
        {
            var workbookPart = document.WorkbookPart
                ?? throw new ReportGenerationException("範本檔案格式不正確，找不到活頁簿內容。");

            var originalSheets = workbookPart.Workbook.Sheets?.Elements<Sheet>().ToList();
            if (originalSheets == null || originalSheets.Count == 0)
            {
                throw new ReportGenerationException("範本檔案格式不正確，找不到工作表。");
            }

            if (originalSheets.Count == 1)
            {
                // 只有一個工作表，沒有其他分頁需要維持原樣，套版結果就是最終結果
                return renderedTargetSheetBytes;
            }

            var targetSheet = ResolveTargetSheet(originalSheets);
            var targetWorksheetPart = (WorksheetPart)workbookPart.GetPartById(targetSheet.Id!.Value!);
            targetWorksheetPart.Worksheet = renderedWorksheet;
            targetWorksheetPart.Worksheet.Save();
        }

        return originalStream.ToArray();
    }

    private static byte[] RemoveDanglingWorkbookRelationships(byte[] xlsxBytes, HashSet<string> relationshipIdsToRemove)
    {
        if (relationshipIdsToRemove.Count == 0) return xlsxBytes;

        using var zipStream = new MemoryStream();
        zipStream.Write(xlsxBytes, 0, xlsxBytes.Length);
        zipStream.Position = 0;

        using (var archive = new ZipArchive(zipStream, ZipArchiveMode.Update))
        {
            var entry = archive.GetEntry("xl/_rels/workbook.xml.rels");
            if (entry == null) return xlsxBytes;

            string content;
            using (var entryStream = entry.Open())
            using (var reader = new StreamReader(entryStream, Encoding.UTF8))
            {
                content = reader.ReadToEnd();
            }

            XNamespace relNs = "http://schemas.openxmlformats.org/package/2006/relationships";
            var relsDoc = XDocument.Parse(content);
            var danglingRelationships = relsDoc.Root!.Elements(relNs + "Relationship")
                .Where(e => relationshipIdsToRemove.Contains(e.Attribute("Id")?.Value ?? ""))
                .ToList();

            if (danglingRelationships.Count == 0) return xlsxBytes;

            foreach (var rel in danglingRelationships)
            {
                rel.Remove();
            }

            entry.Delete();
            var newEntry = archive.CreateEntry("xl/_rels/workbook.xml.rels");
            using var writeStream = newEntry.Open();
            using var writer = new StreamWriter(writeStream, new UTF8Encoding(false));
            writer.Write(relsDoc.Declaration + relsDoc.ToString(SaveOptions.DisableFormatting));
        }

        return zipStream.ToArray();
    }

    /// <summary>移除範本中殘留、指向被移除工作表的具名範圍（definedName），避免留下無效參照。
    /// 若具名範圍本來就是綁定在保留下來的目標工作表，重新編號成 0（因為裁切後只剩它一個）。</summary>
    private static void RemoveDanglingDefinedNames(DefinedNames? definedNames, uint targetOriginalIndex, HashSet<string> removedSheetNames)
    {
        if (definedNames == null) return;

        var toRemove = new List<DefinedName>();
        foreach (var name in definedNames.Elements<DefinedName>())
        {
            if (name.LocalSheetId != null)
            {
                if (name.LocalSheetId.Value == targetOriginalIndex)
                {
                    name.LocalSheetId = 0;
                }
                else
                {
                    toRemove.Add(name);
                }
            }
            else if (ReferencesAnySheet(name.Text, removedSheetNames))
            {
                toRemove.Add(name);
            }
        }

        foreach (var name in toRemove)
        {
            name.Remove();
        }

        if (!definedNames.Elements<DefinedName>().Any())
        {
            definedNames.Remove();
        }
    }

    private static bool ReferencesAnySheet(string? formulaText, HashSet<string> sheetNames)
    {
        if (string.IsNullOrEmpty(formulaText)) return false;
        foreach (var name in sheetNames)
        {
            if (formulaText.StartsWith(name + "!", StringComparison.Ordinal) ||
                formulaText.StartsWith($"'{name}'!", StringComparison.Ordinal))
            {
                return true;
            }
        }
        return false;
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
