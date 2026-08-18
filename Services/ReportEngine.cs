using System.Diagnostics;
using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using excelReport.Models;
using Microsoft.Extensions.Caching.Memory;
using MiniExcelLibs;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using A = DocumentFormat.OpenXml.Drawing;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace excelReport.Services;

public class ReportEngine : IReportEngine
{
    private readonly IReportConfigStore _configStore;
    private readonly IDataSourceClient _dataSourceClient;
    private readonly IReportTemplateScanner _scanner;
    private readonly IHttpClientFactory _httpClientFactory;
    private readonly IWebHostEnvironment _env;
    private readonly IMemoryCache _cache;
    private readonly ILogger<ReportEngine> _logger;

    public ReportEngine(
        IReportConfigStore configStore,
        IDataSourceClient dataSourceClient,
        IReportTemplateScanner scanner,
        IHttpClientFactory httpClientFactory,
        IWebHostEnvironment env,
        IMemoryCache cache,
        ILogger<ReportEngine> logger)
    {
        _configStore = configStore;
        _dataSourceClient = dataSourceClient;
        _scanner = scanner;
        _httpClientFactory = httpClientFactory;
        _env = env;
        _cache = cache;
        _logger = logger;
    }

    public async Task<ReportGenerationResult> GenerateAsync(string code, IDictionary<string, string> parameters, CancellationToken cancellationToken = default)
    {
        var stopwatch = Stopwatch.StartNew();
        ReportConfig? config = null;
        try
        {
            config = _configStore.Get(code)
                ?? throw new ReportGenerationException($"找不到報表設定：{code}");

            ValidateParameters(config, parameters);

            // 取得要填入的資料
            var json = await _dataSourceClient.FetchAsync(config.DataSource, parameters, cancellationToken);

            JObject root;
            try
            {
                // 不用 JObject.Parse(json)：其預設 DateParseHandling 會把長得像日期的字串（例如
                // ISO 8601 時間戳）自動解析成 DateTime，導致 TokenToValue 用伺服器當地文化格式
                // （例如 zh-TW 的「2026/7/3 下午 01:34:29」）覆蓋掉原始文字。改用 JsonTextReader
                // 並關閉 DateParseHandling，讓所有字串節點維持來源 API 回傳的原始文字。
                using var stringReader = new StringReader(json);
                using var jsonReader = new JsonTextReader(stringReader) { DateParseHandling = DateParseHandling.None };
                if (JToken.ReadFrom(jsonReader) is not JObject parsed)
                {
                    throw new ReportGenerationException("資料來源回傳的內容不是 JSON 物件。");
                }
                root = parsed;
            }
            catch (JsonException ex)
            {
                throw new ReportGenerationException($"資料來源回傳的內容不是有效的 JSON：{ex.Message}", ex);
            }

            var templateBytes = await LoadTemplateBytesAsync(config.TemplateFile);
            // 讀取excel sheet
            var scanResult = _scanner.Scan(templateBytes, config.TemplateSheet);

            var rootToken = ResolveRoot(root, config.Mapping.Root);
            var imageValues = ResolveImages(config.Mapping, rootToken);

            // 用範本實際掃到的標記驅動資料字典的建立（而非只看設定裡填了什麼），確保範本裡
            // 每個標記都會有對應的 key——即使某個標記完全沒有在映射設定填寫任何內容也一樣。
            // 少了這一步，MiniExcel 套版時遇到資料字典裡不存在的 key 會拋出 NullReferenceException。
            var data = BuildDataDictionary(config.Mapping, scanResult, rootToken);
            var fileBytes = await RenderTemplateAsync(templateBytes, data, cancellationToken);

            if (config.SubReports.Count > 0)
            {
                // 子報表：各自獨立的 xlsx 範本，依序渲染、接在主報表（以及前一個子報表）後面。
                fileBytes = await AppendSubReportsAsync(config.SubReports, fileBytes, config.TemplateSheet, rootToken, cancellationToken);
            }

            if (config.Mapping.Images.Count > 0)
            {
                fileBytes = await EmbedImagesAsync(fileBytes, config.Mapping.Images, imageValues, cancellationToken);
            }
            var fileName = BuildFileName(config, parameters);

            stopwatch.Stop();
            WriteLog(config.Code, parameters, stopwatch.ElapsedMilliseconds, true, null);

            return new ReportGenerationResult { FileBytes = fileBytes, FileName = fileName };
        }
        catch (ReportGenerationException ex)
        {
            stopwatch.Stop();
            WriteLog(config?.Code ?? code, parameters, stopwatch.ElapsedMilliseconds, false, ex.Message);
            throw;
        }
        catch (Exception ex)
        {
            stopwatch.Stop();
            _logger.LogError(ex, "產生報表 {Code} 時發生未預期錯誤", code);
            WriteLog(config?.Code ?? code, parameters, stopwatch.ElapsedMilliseconds, false, ex.Message);
            throw new ReportGenerationException($"產生報表時發生未預期的錯誤：{ex.Message}", ex);
        }
    }

    private static void ValidateParameters(ReportConfig config, IDictionary<string, string> parameters)
    {
        var missing = config.Parameters
            .Where(p => p.Required && (!parameters.TryGetValue(p.Key, out var v) || string.IsNullOrWhiteSpace(v)))
            .Select(p => p.Label)
            .ToList();

        if (missing.Count > 0)
        {
            throw new ReportGenerationException($"缺少必填參數：{string.Join("、", missing)}");
        }
    }

    private static JToken ResolveRoot(JObject root, string rootPath)
    {
        if (string.IsNullOrWhiteSpace(rootPath) || rootPath == "$")
        {
            return root;
        }

        var token = root.SelectToken(rootPath);
        if (token == null)
        {
            throw new ReportGenerationException($"找不到資料根節點，JSONPath：{rootPath}");
        }
        return token;
    }

    /// <summary>
    /// 保留欄位名，永遠對應空字串。範本中每個單值標記後面都會附加 {{_str}}（或集合內的
    /// {{collection._str}}），讓儲存格內含多個標記，藉此讓 MiniExcel 套版時保留文字型別，
    /// 避免像 "0040" 這種數字外觀的字串被誤判為數值而遺失前導零。
    /// </summary>
    private const string ReservedColumnName = "_str";

    /// <summary>範本標記沒有填寫對應 JSONPath 時的預設路徑："$.{標記名稱}"。</summary>
    private static string EffectivePath(string fieldKey, string? configuredPath) =>
        string.IsNullOrWhiteSpace(configuredPath) ? $"$.{fieldKey}" : configuredPath;

    /// <summary>只建立單值欄位（fields）那一份資料，供一般路徑與巢狀子表路徑共用。</summary>
    private static Dictionary<string, object> BuildFieldsData(MappingConfig mapping, TemplateScanResult scanResult, JToken rootToken)
    {
        var data = new Dictionary<string, object> { [ReservedColumnName] = "" };

        foreach (var fieldKey in scanResult.Fields)
        {
            mapping.Fields.TryGetValue(fieldKey, out var configuredPath);
            var token = rootToken.SelectToken(EffectivePath(fieldKey, configuredPath));
            data[fieldKey] = token == null ? "" : TokenToValue(token);
        }

        return data;
    }

    /// <summary>
    /// 以範本實際掃到的標記（scanResult）為準來建立資料字典，而不是只看 mapping 設定裡有寫什麼。
    /// mapping 只提供「有填寫的話要用哪個 JSONPath」
    /// 所有標記都會得到一個對應的資料字典 key（找不到節點就給空字串）。
    /// </summary>
    private static Dictionary<string, object> BuildDataDictionary(MappingConfig mapping, TemplateScanResult scanResult, JToken rootToken)
    {
        var data = BuildFieldsData(mapping, scanResult, rootToken);

        foreach (var collectionScan in scanResult.Collections)
        {
            var collectionMapping = mapping.Collections.FirstOrDefault(c => c.Name == collectionScan.Name);
            data[collectionScan.Name] = BuildCollectionRows(
                rootToken, collectionScan.Name, collectionMapping?.Path, collectionScan.Columns,
                collectionMapping?.Columns ?? new Dictionary<string, string>(), collectionMapping?.Spread);
        }

        return data;
    }

    /// <summary>解析一個集合節點（相對 scopeToken）並依範本掃到的欄位建立資料列。用 SelectTokens
    /// 依標準 JSONPath 語意解析路徑，同時接受兩種寫法：「$.items」這種指向陣列容器本身的路徑
    /// （只查到一個 token、且該 token 是陣列）會展開其元素；「$.items[*]」這種帶萬用字元、直接
    /// 查到多個元素的路徑，每個查到的 token 就各自是一列。找不到對應節點、或節點是純量卻寫成
    /// 點記法，視為沒有明細列，不中斷報表產生。主報表 Collections 跟子報表 Collections 都共用
    /// 這條路徑，行為完全一致；跟 ApplySpread 解析 From 用的是同一套「容器 vs 已展開」判斷邏輯。</summary>
    private static List<IDictionary<string, object>> BuildCollectionRows(
        JToken scopeToken, string name, string? configuredPath, List<string> columns,
        Dictionary<string, string> configuredColumns, SpreadConfig? spread)
    {
        var tokens = scopeToken.SelectTokens(EffectivePath(name, configuredPath)).ToList();
        var items = tokens.Count == 1 && tokens[0].Type == JTokenType.Array
            ? tokens[0].ToList()
            : tokens;

        var rows = new List<IDictionary<string, object>>();
        var index = 0;
        foreach (var item in items)
        {
            rows.Add(BuildRowFromItem(columns, configuredColumns, item, index, spread));
            index++;
        }

        return rows;
    }

    private static void ApplySpread(JToken item, SpreadConfig spread, Dictionary<string, object> row)
    {
        // SelectTokens（複數）同時涵蓋兩種來源寫法：
        // - "$.values" 這種指向單一陣列節點的路徑，會回傳「一個」token（該陣列本身），此時展開其元素；
        // - "$.samples[*].measuredValue" 這種帶萬用字元、跨多個節點取值的路徑，會直接回傳「多個」純量 token。
        var tokens = item.SelectTokens(spread.From).ToList();
        var elements = tokens.Count == 1 && tokens[0].Type == JTokenType.Array
            ? tokens[0].ToList()
            : tokens;

        // Property 有填寫時，陣列元素視為物件，用標準 JSONPath 語法解析（root 是每個元素本身）
        // 取其中的值當展開值；留空時維持舊行為，直接用元素本身的值。
        var values = string.IsNullOrWhiteSpace(spread.Property)
            ? elements.Select(TokenToValue).ToList()
            : elements.Select(e => e.SelectToken(spread.Property) is { } t ? TokenToValue(t) : "").ToList();

        for (var i = 0; i < spread.Max; i++)
        {
            var key = $"{spread.Prefix}{i + 1}";
            row[key] = i < values.Count ? values[i] : "";
        }
    }

    private static string TokenToValue(JToken token)
    {
        if (token.Type == JTokenType.Null) return "";
        if (token is JValue jv) return jv.Value?.ToString() ?? "";
        return token.ToString();
    }

    /// <summary>集合列欄位的保留字：自動填入該元素在陣列中的序號（從 1 開始），不走 JSONPath。</summary>
    private const string SeqColumnName = "seq";

    /// <summary>依範本掃到的欄位清單，用一個已知的 item（陣列元素或單一物件）建立一列資料。
    /// columns 是範本實際掃到的欄位名稱清單，configuredColumns 是設定檔裡「有填寫的話要用哪個
    /// JSONPath」。主報表 Collections、子報表父層列、子報表子層集合都共用這個建列邏輯。</summary>
    private static Dictionary<string, object> BuildRowFromItem(
        List<string> columns, Dictionary<string, string> configuredColumns, JToken item, int index, SpreadConfig? spread)
    {
        var row = new Dictionary<string, object> { [ReservedColumnName] = "" };
        foreach (var columnKey in columns)
        {
            if (columnKey == SeqColumnName && !configuredColumns.ContainsKey(SeqColumnName))
            {
                row[columnKey] = (index + 1).ToString(CultureInfo.InvariantCulture);
                continue;
            }

            configuredColumns.TryGetValue(columnKey, out var configuredPath);
            var valueToken = item.SelectToken(EffectivePath(columnKey, configuredPath));
            row[columnKey] = valueToken == null ? "" : TokenToValue(valueToken);
        }

        if (spread != null)
        {
            ApplySpread(item, spread, row);
        }

        return row;
    }

    /// <summary>
    /// 依序渲染每個子報表、接在 mainBytes（主報表，或前一個子報表接完後的結果）後面。每個子報表
    /// 都是獨立的 xlsx 範本檔，綁定一個根節點 JSONPath：用 SelectTokens 依標準 JSONPath 語意查到
    /// 幾個節點，就對每個節點各自渲染一次整份子範本——該次渲染的 Fields/Collections/Images 都
    /// 相對這個節點解析，直接重用主報表在用的 BuildDataDictionary/ResolveImages，用法完全比照
    /// 主報表。再把每次渲染結果的列（含內嵌圖片）取出來、往下平移列號後接到目前輸出的工作表
    /// 後面。
    /// </summary>
    private async Task<byte[]> AppendSubReportsAsync(
        List<SubReportConfig> subReports, byte[] mainBytes, string mainSheet, JToken rootToken, CancellationToken cancellationToken)
    {
        using var accumulatorStream = new MemoryStream();
        accumulatorStream.Write(mainBytes, 0, mainBytes.Length);
        accumulatorStream.Position = 0;

        using (var accumulatorDoc = SpreadsheetDocument.Open(accumulatorStream, true))
        {
            var currentMaxRow = GetMaxRowIndex(accumulatorDoc, _scanner, mainSheet);

            foreach (var subReport in subReports)
            {
                if (string.IsNullOrWhiteSpace(subReport.TemplateFile)) continue;

                var subTemplateBytes = await LoadTemplateBytesAsync(subReport.TemplateFile);
                var subScanResult = _scanner.Scan(subTemplateBytes, subReport.TemplateSheet);

                foreach (var renderRoot in ResolveSubReportRoots(rootToken, subReport.Root))
                {
                    var data = BuildDataDictionary(subReport.Mapping, subScanResult, renderRoot);
                    var renderedBytes = await RenderTemplateAsync(subTemplateBytes, data, cancellationToken);

                    if (subReport.Mapping.Images.Count > 0)
                    {
                        var imageValues = ResolveImages(subReport.Mapping, renderRoot);
                        renderedBytes = await EmbedImagesAsync(renderedBytes, subReport.Mapping.Images, imageValues, cancellationToken);
                    }

                    currentMaxRow = AppendAllRows(accumulatorDoc, _scanner, mainSheet, renderedBytes, subReport.TemplateSheet, currentMaxRow);
                }
            }
        }

        return accumulatorStream.ToArray();
    }

    /// <summary>依標準 JSONPath 語意解析子報表的根節點：查到幾個節點就回傳幾個，每個節點各對應
    /// 一次渲染。帶萬用字元的路徑（如 "$.tasks[*]"）語意上查到多個節點；不帶萬用字元的路徑
    /// （如 "$"、"$.data"、"$.records[0]"）語意上查到單一節點——都交給 SelectTokens 本身的
    /// JSONPath 語意判斷，不做任何額外的字串層面判斷。路徑留空時視為 "$"，直接回傳 rootToken
    /// 本身。</summary>
    private static IEnumerable<JToken> ResolveSubReportRoots(JToken rootToken, string rootPath)
    {
        if (string.IsNullOrWhiteSpace(rootPath))
        {
            yield return rootToken;
            yield break;
        }

        foreach (var token in rootToken.SelectTokens(rootPath))
        {
            yield return token;
        }
    }

    private static int GetMaxRowIndex(SpreadsheetDocument document, IReportTemplateScanner scanner, string? sheet)
    {
        var worksheetPart = scanner.GetTargetWorksheetPart(document.WorkbookPart!, sheet);
        var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
        var rows = sheetData?.Elements<Row>().ToList() ?? new List<Row>();
        return rows.Count == 0 ? 0 : (int)rows.Max(r => r.RowIndex?.Value ?? 0);
    }

    /// <summary>把 sourceBytes（某個子報表的完整渲染結果）的所有列（含合併儲存格範圍），依
    /// offset 平移列號後接到 accumulatorDoc 目前的工作表後面。子報表是獨立的 xlsx 檔，不像
    /// 主報表混雜固定欄位，所以整份檔案的每一列都屬於要接上去的內容，不需要另外切出範圍。
    /// accumulatorSheet 是主報表自己選的工作表(整個拼接過程中固定不變)，sourceSheet 是這個
    /// 子報表自己選的工作表，兩者可能不同。主報表與子報表用各自獨立的範本渲染，styles.xml/
    /// sharedStrings 彼此獨立；文字一律轉成 inlineStr，避免要處理「兩邊各自獨立的 sharedStrings
    /// 索引對不上」的問題。</summary>
    private static int AppendAllRows(
        SpreadsheetDocument accumulatorDoc, IReportTemplateScanner scanner, string? accumulatorSheet,
        byte[] sourceBytes, string? sourceSheet, int currentMaxRow)
    {
        var accumulatorSheetPart = scanner.GetTargetWorksheetPart(accumulatorDoc.WorkbookPart!, accumulatorSheet);
        var accumulatorSheetData = accumulatorSheetPart.Worksheet.GetFirstChild<SheetData>()
            ?? accumulatorSheetPart.Worksheet.AppendChild(new SheetData());

        using var sourceStream = new MemoryStream(sourceBytes);
        using var sourceDoc = SpreadsheetDocument.Open(sourceStream, false);
        var sourceWorkbookPart = sourceDoc.WorkbookPart!;
        var sourceSheetPart = scanner.GetTargetWorksheetPart(sourceWorkbookPart, sourceSheet);
        var sourceSharedStrings = sourceWorkbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
        var sourceSheetData = sourceSheetPart.Worksheet.GetFirstChild<SheetData>();

        var rowsToCopy = (sourceSheetData?.Elements<Row>() ?? Enumerable.Empty<Row>())
            .Where(r => r.RowIndex is not null)
            .OrderBy(r => r.RowIndex!.Value)
            .ToList();

        var offset = currentMaxRow;

        foreach (var srcRow in rowsToCopy)
        {
            var newRowIndex = (uint)(srcRow.RowIndex!.Value + offset);
            var newRow = new Row { RowIndex = newRowIndex };
            if (srcRow.Height is not null) newRow.Height = srcRow.Height.Value;
            if (srcRow.CustomHeight is not null) newRow.CustomHeight = srcRow.CustomHeight.Value;
            if (srcRow.ThickBot is not null) newRow.ThickBot = srcRow.ThickBot.Value;

            foreach (var srcCell in srcRow.Elements<Cell>())
            {
                var newCell = new Cell
                {
                    CellReference = ShiftCellReferenceRow(srcCell.CellReference!.Value!, offset),
                    StyleIndex = srcCell.StyleIndex
                };

                if (srcCell.DataType?.Value == CellValues.SharedString)
                {
                    var text = ResolveSharedStringText(srcCell, sourceSharedStrings);
                    if (!string.IsNullOrEmpty(text))
                    {
                        newCell.DataType = CellValues.InlineString;
                        newCell.InlineString = new InlineString(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
                    }
                }
                else if (srcCell.DataType?.Value == CellValues.InlineString)
                {
                    if (srcCell.InlineString is not null)
                    {
                        newCell.DataType = CellValues.InlineString;
                        newCell.InlineString = (InlineString)srcCell.InlineString.CloneNode(true);
                    }
                }
                else
                {
                    if (srcCell.DataType is not null) newCell.DataType = srcCell.DataType.Value;
                    if (srcCell.CellValue is not null) newCell.CellValue = new CellValue(srcCell.CellValue.Text);
                }

                newRow.Append(newCell);
            }

            accumulatorSheetData.Append(newRow);
        }

        var sourceMergeCells = sourceSheetPart.Worksheet.GetFirstChild<MergeCells>();
        if (sourceMergeCells is not null)
        {
            var accumulatorMergeCells = accumulatorSheetPart.Worksheet.GetFirstChild<MergeCells>();
            if (accumulatorMergeCells is null)
            {
                accumulatorMergeCells = new MergeCells();
                accumulatorSheetPart.Worksheet.InsertAfter(accumulatorMergeCells, accumulatorSheetData);
            }

            foreach (var mc in sourceMergeCells.Elements<MergeCell>())
            {
                var (startRef, endRef) = SplitMergeRange(mc.Reference!.Value!);
                var (startCol, startRow) = ParseCellReferenceParts(startRef);
                var (endCol, endRow) = ParseCellReferenceParts(endRef);
                accumulatorMergeCells.Append(new MergeCell { Reference = $"{startCol}{startRow + offset}:{endCol}{endRow + offset}" });
            }

            accumulatorMergeCells.Count = (uint)accumulatorMergeCells.Elements<MergeCell>().Count();
        }

        CopyDrawings(accumulatorSheetPart, sourceSheetPart, offset);

        var newMaxRow = rowsToCopy.Count > 0 ? (int)(rowsToCopy.Max(r => r.RowIndex!.Value) + offset) : currentMaxRow;

        var dimension = accumulatorSheetPart.Worksheet.GetFirstChild<SheetDimension>();
        if (dimension is not null)
        {
            var (startRef, _) = SplitMergeRange(dimension.Reference!.Value!);
            var (startCol, startRow) = ParseCellReferenceParts(startRef);
            dimension.Reference = $"{startCol}{startRow}:XFD{newMaxRow}";
        }

        accumulatorSheetPart.Worksheet.Save();
        return newMaxRow;
    }

    /// <summary>把 sourceSheetPart 裡已內嵌的圖片（子報表 Mapping.Images 對單次渲染出的獨立檔案
    /// 呼叫 EmbedImagesAsync 內嵌好的圖片），依 offset 平移列號後複製到 accumulatorSheetPart。
    /// xdr 的列/欄索引本身就是 0-based 的位移量，跟 Row/MergeCell 用的 1-based 列號一樣，直接
    /// 加上 offset 位移即可，不需要額外轉換。</summary>
    private static void CopyDrawings(WorksheetPart accumulatorSheetPart, WorksheetPart sourceSheetPart, int offset)
    {
        var sourceDrawingsPart = sourceSheetPart.DrawingsPart;
        if (sourceDrawingsPart?.WorksheetDrawing == null) return;

        foreach (var srcAnchor in sourceDrawingsPart.WorksheetDrawing.Elements<Xdr.OneCellAnchor>())
        {
            var blip = srcAnchor.Descendants<A.Blip>().FirstOrDefault();
            var fromMarker = srcAnchor.FromMarker;
            var extent = srcAnchor.Extent;
            if (blip?.Embed?.Value == null || fromMarker?.ColumnId == null || fromMarker.RowId == null || extent == null) continue;

            var srcImagePart = (ImagePart)sourceDrawingsPart.GetPartById(blip.Embed.Value);
            using var imageStream = srcImagePart.GetStream();
            using var buffer = new MemoryStream();
            imageStream.CopyTo(buffer);
            var imageBytes = buffer.ToArray();

            var col = int.Parse(fromMarker.ColumnId.Text, CultureInfo.InvariantCulture);
            var row = int.Parse(fromMarker.RowId.Text, CultureInfo.InvariantCulture) + offset;

            InsertPicture(
                accumulatorSheetPart, col, row, imageBytes, DetectImagePartType(imageBytes),
                extent.Cx?.Value ?? ImageWidthEmu, extent.Cy?.Value ?? ImageHeightEmu);
        }
    }

    private static string ResolveSharedStringText(Cell cell, SharedStringTablePart? sharedStrings)
    {
        if (sharedStrings == null || cell.CellValue == null) return "";
        if (!int.TryParse(cell.CellValue.Text, out var idx)) return "";
        var items = sharedStrings.SharedStringTable.Elements<SharedStringItem>().ToList();
        return idx >= 0 && idx < items.Count ? items[idx].InnerText : "";
    }

    private static (string Start, string End) SplitMergeRange(string reference)
    {
        var parts = reference.Split(':');
        return parts.Length == 2 ? (parts[0], parts[1]) : (parts[0], parts[0]);
    }

    private static (string Column, int Row) ParseCellReferenceParts(string cellReference)
    {
        var letters = new string(cellReference.TakeWhile(char.IsLetter).ToArray());
        var digits = cellReference[letters.Length..];
        return (letters, int.Parse(digits, CultureInfo.InvariantCulture));
    }

    private static string ShiftCellReferenceRow(string cellReference, int offset)
    {
        var (col, row) = ParseCellReferenceParts(cellReference);
        return $"{col}{row + offset}";
    }

    /// <summary>解析圖片欄位的原始值（URL 或 base64 內容）。path 以 "$" 開頭視為 JSONPath；
    /// 否則直接當成固定值使用（例如寫死的圖片 URL），不透過 JSONPath 解析。
    /// 找不到節點或值為空的欄位直接略過，不會出現在回傳結果中——EmbedImagesAsync 之後
    /// 只會處理有值的圖片標記，其餘標記清空文字即可。</summary>
    private static Dictionary<string, string> ResolveImages(MappingConfig mapping, JToken rootToken)
    {
        var result = new Dictionary<string, string>();
        foreach (var (fieldKey, imageMapping) in mapping.Images)
        {
            if (string.IsNullOrWhiteSpace(imageMapping.Path)) continue;

            string value;
            if (imageMapping.Path.StartsWith('$'))
            {
                var token = rootToken.SelectToken(imageMapping.Path);
                if (token == null || token.Type == JTokenType.Null) continue;
                value = TokenToValue(token);
            }
            else
            {
                value = imageMapping.Path;
            }

            if (!string.IsNullOrWhiteSpace(value))
            {
                result[fieldKey] = value;
            }
        }
        return result;
    }

    private static readonly System.Text.RegularExpressions.Regex ImageMarkerRegex =
        new(@"\[\[\s*([a-zA-Z0-9_.]+)\s*\]\]", System.Text.RegularExpressions.RegexOptions.Compiled);

    /// <summary>
    /// MiniExcel 完全不認得 [[fieldName]] 這種語法（它只處理 {{}}），套版後會原樣留在儲存格裡。
    /// 這一步在套版完成後，另外用 OpenXML 掃描含 [[]] 標記的儲存格：清空標記文字，並在有對應
    /// 圖片值時把圖片（下載或 base64 解碼）用 OneCellAnchor 嵌入到該儲存格位置。
    /// 找不到圖片值、下載/解碼失敗都只會略過該張圖片，不會讓整份報表產生失敗。
    /// </summary>
    private async Task<byte[]> EmbedImagesAsync(
        byte[] fileBytes,
        Dictionary<string, ImageMapping> imageMappings,
        Dictionary<string, string> imageValues,
        CancellationToken cancellationToken)
    {
        using var stream = new MemoryStream();
        stream.Write(fileBytes, 0, fileBytes.Length);
        stream.Position = 0;

        using (var document = SpreadsheetDocument.Open(stream, true))
        {
            var workbookPart = document.WorkbookPart;
            if (workbookPart == null) return fileBytes;

            var sharedStrings = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();

            foreach (var worksheetPart in workbookPart.WorksheetParts)
            {
                var worksheetChanged = false;
                var cells = worksheetPart.Worksheet.Descendants<Cell>().ToList();

                foreach (var cell in cells)
                {
                    var text = GetCellText(cell, sharedStrings);
                    if (string.IsNullOrEmpty(text)) continue;

                    var match = ImageMarkerRegex.Match(text);
                    if (!match.Success) continue;

                    var fieldName = match.Groups[1].Value;
                    ClearCellText(cell);
                    worksheetChanged = true;

                    if (!imageValues.TryGetValue(fieldName, out var sourceValue) || cell.CellReference?.Value == null)
                    {
                        continue;
                    }

                    var sourceType = imageMappings.TryGetValue(fieldName, out var imageMapping) ? imageMapping.SourceType : "auto";

                    byte[]? imageBytes;
                    try
                    {
                        imageBytes = await ResolveImageBytesAsync(sourceValue, sourceType, cancellationToken);
                    }
                    catch (Exception ex)
                    {
                        _logger.LogWarning(ex, "圖片欄位 {Field} 下載/解碼失敗，略過嵌入", fieldName);
                        imageBytes = null;
                    }

                    if (imageBytes is { Length: > 0 })
                    {
                        InsertPicture(worksheetPart, cell.CellReference.Value, imageBytes);
                    }
                }

                if (worksheetChanged)
                {
                    worksheetPart.Worksheet.Save();
                }
            }
        }

        return stream.ToArray();
    }

    private static void ClearCellText(Cell cell)
    {
        cell.CellValue = new CellValue("");
        cell.DataType = CellValues.String;
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

    private async Task<byte[]?> ResolveImageBytesAsync(string sourceValue, string sourceType, CancellationToken cancellationToken)
    {
        var isUrl = sourceType == "url" || (sourceType != "base64" &&
            (sourceValue.StartsWith("http://", StringComparison.OrdinalIgnoreCase) ||
             sourceValue.StartsWith("https://", StringComparison.OrdinalIgnoreCase)));

        if (isUrl)
        {
            using var client = _httpClientFactory.CreateClient("ReportDataSource");
            client.Timeout = TimeSpan.FromSeconds(30);
            return await client.GetByteArrayAsync(sourceValue, cancellationToken);
        }

        var base64 = sourceValue;
        var commaIndex = base64.IndexOf(',');
        if (base64.StartsWith("data:", StringComparison.OrdinalIgnoreCase) && commaIndex >= 0)
        {
            base64 = base64[(commaIndex + 1)..];
        }
        return Convert.FromBase64String(base64);
    }

    // 圖片固定尺寸（EMU，1px = 9525 EMU）：約 120x70px，配合表頭列高與右上角保留的 4 欄寬度。
    private const long ImageWidthEmu = 1143000;
    private const long ImageHeightEmu = 666750;

    private static void InsertPicture(WorksheetPart worksheetPart, string cellReference, byte[] imageBytes)
    {
        var (col, row) = ParseCellReference(cellReference);
        InsertPicture(worksheetPart, col, row, imageBytes, DetectImagePartType(imageBytes), ImageWidthEmu, ImageHeightEmu);
    }

    /// <summary>把圖片以 OneCellAnchor 插入到工作表指定的 0-based (col,row) 位置。子報表列拼接
    /// （CopyDrawings）跟一般的 [[field]] 圖片標記嵌入（上面的 cellReference 版本）共用這個核心
    /// 邏輯，差別只在圖片來源尺寸/內容類型是否要沿用原圖，還是用固定的預設值。</summary>
    private static void InsertPicture(
        WorksheetPart worksheetPart, int col, int row, byte[] imageBytes, PartTypeInfo contentType, long widthEmu, long heightEmu)
    {
        var drawingsPart = worksheetPart.DrawingsPart;
        Xdr.WorksheetDrawing worksheetDrawing;
        if (drawingsPart == null)
        {
            drawingsPart = worksheetPart.AddNewPart<DrawingsPart>();
            worksheetDrawing = new Xdr.WorksheetDrawing();
            drawingsPart.WorksheetDrawing = worksheetDrawing;
            worksheetPart.Worksheet.Append(new Drawing { Id = worksheetPart.GetIdOfPart(drawingsPart) });
        }
        else
        {
            worksheetDrawing = drawingsPart.WorksheetDrawing;
        }

        var imagePart = drawingsPart.AddImagePart(contentType);
        using (var imageStream = new MemoryStream(imageBytes))
        {
            imagePart.FeedData(imageStream);
        }
        var imageRelId = drawingsPart.GetIdOfPart(imagePart);

        var picId = (uint)(worksheetDrawing.Elements<Xdr.OneCellAnchor>().Count() + 1);

        var anchor = new Xdr.OneCellAnchor(
            new Xdr.FromMarker
            {
                ColumnId = new Xdr.ColumnId(col.ToString()),
                ColumnOffset = new Xdr.ColumnOffset("0"),
                RowId = new Xdr.RowId(row.ToString()),
                RowOffset = new Xdr.RowOffset("0")
            },
            new Xdr.Extent { Cx = widthEmu, Cy = heightEmu },
            new Xdr.Picture(
                new Xdr.NonVisualPictureProperties(
                    new Xdr.NonVisualDrawingProperties { Id = picId, Name = $"Picture {picId}" },
                    new Xdr.NonVisualPictureDrawingProperties(new A.PictureLocks { NoChangeAspect = true })),
                new Xdr.BlipFill(
                    new A.Blip { Embed = imageRelId },
                    new A.Stretch(new A.FillRectangle())),
                new Xdr.ShapeProperties(
                    new A.Transform2D(
                        new A.Offset { X = 0, Y = 0 },
                        new A.Extents { Cx = widthEmu, Cy = heightEmu }),
                    new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle })),
            new Xdr.ClientData());

        worksheetDrawing.Append(anchor);
        worksheetDrawing.Save();
    }

    private static (int Col, int Row) ParseCellReference(string cellReference)
    {
        var letters = new string(cellReference.TakeWhile(char.IsLetter).ToArray());
        var digits = cellReference[letters.Length..];
        var col = 0;
        foreach (var c in letters)
        {
            col = col * 26 + (char.ToUpperInvariant(c) - 'A' + 1);
        }
        var row = int.Parse(digits, CultureInfo.InvariantCulture);
        return (col - 1, row - 1); // 轉成 0-based，符合 xdr:from 的欄/列索引
    }

    private static PartTypeInfo DetectImagePartType(byte[] bytes)
    {
        if (bytes.Length >= 8 && bytes[0] == 0x89 && bytes[1] == 0x50 && bytes[2] == 0x4E && bytes[3] == 0x47)
        {
            return ImagePartType.Png;
        }
        if (bytes.Length >= 3 && bytes[0] == 0xFF && bytes[1] == 0xD8 && bytes[2] == 0xFF)
        {
            return ImagePartType.Jpeg;
        }
        if (bytes.Length >= 6 && bytes[0] == 'G' && bytes[1] == 'I' && bytes[2] == 'F')
        {
            return ImagePartType.Gif;
        }
        return ImagePartType.Png;
    }

    private async Task<byte[]> LoadTemplateBytesAsync(string templateFile)
    {
        var templatePath = Path.Combine(_env.ContentRootPath, "App_Data", "templates", templateFile);
        if (!File.Exists(templatePath))
        {
            throw new ReportGenerationException($"找不到範本檔案：{templateFile}");
        }

        var lastWriteTicks = File.GetLastWriteTimeUtc(templatePath).Ticks;
        var cacheKey = $"tpl:{templateFile}:{lastWriteTicks}";
        return await _cache.GetOrCreateAsync(cacheKey, entry =>
        {
            entry.SlidingExpiration = TimeSpan.FromMinutes(30);
            return Task.FromResult(File.ReadAllBytes(templatePath));
        }) ?? throw new ReportGenerationException("範本檔案讀取失敗。");
    }

    private static async Task<byte[]> RenderTemplateAsync(byte[] templateBytes, Dictionary<string, object> data, CancellationToken cancellationToken)
    {
        using var outputStream = new MemoryStream();
        try
        {
            await outputStream.SaveAsByTemplateAsync(templateBytes, data, cancellationToken: cancellationToken);
        }
        catch (Exception ex) when (ex is not ReportGenerationException)
        {
            throw new ReportGenerationException($"套版產生 xlsx 時發生錯誤：{ex.Message}", ex);
        }

        return outputStream.ToArray();
    }

    private static string BuildFileName(ReportConfig config, IDictionary<string, string> parameters)
    {
        var keyParam = parameters.TryGetValue("docNo", out var docNo) && !string.IsNullOrWhiteSpace(docNo)
            ? docNo
            : parameters.Values.FirstOrDefault(v => !string.IsNullOrWhiteSpace(v)) ?? "";

        var namePart = string.IsNullOrWhiteSpace(keyParam) ? config.Name : $"{config.Name}_{keyParam}";
        return $"{namePart}_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
    }

    private void WriteLog(string code, IDictionary<string, string> parameters, long elapsedMs, bool success, string? error)
    {
        try
        {
            var logsDir = Path.Combine(_env.ContentRootPath, "App_Data", "logs");
            Directory.CreateDirectory(logsDir);
            var logFile = Path.Combine(logsDir, $"report-{DateTime.Now:yyyyMMdd}.log");

            var entry = new
            {
                time = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
                code,
                parameters,
                elapsedMs,
                success,
                error
            };

            File.AppendAllText(logFile, JsonConvert.SerializeObject(entry) + Environment.NewLine);
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "寫入報表產生記錄檔失敗");
        }
    }
}
