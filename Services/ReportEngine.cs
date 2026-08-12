using System.Diagnostics;
using System.Globalization;
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

            var templateBytes = await LoadTemplateBytesAsync(config);
            var scanResult = _scanner.Scan(templateBytes);

            var rootToken = ResolveRoot(root, config.Mapping.Root);
            // 用範本實際掃到的標記驅動資料字典的建立（而非只看設定裡填了什麼），確保範本裡
            // 每個標記都會有對應的 key——即使某個標記完全沒有在映射設定填寫任何內容也一樣。
            // 少了這一步，MiniExcel 套版時遇到資料字典裡不存在的 key 會拋出 NullReferenceException。
            var data = BuildDataDictionary(config.Mapping, scanResult, rootToken);
            var imageValues = ResolveImages(config.Mapping, rootToken);

            var fileBytes = await RenderTemplateAsync(templateBytes, data, cancellationToken);
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

    /// <summary>
    /// 以範本實際掃到的標記（scanResult）為準來建立資料字典，而不是只看 mapping 設定裡有寫什麼。
    /// mapping 只提供「有填寫的話要用哪個 JSONPath」；範本裡任何標記，無論設定裡有沒有寫、寫的是
    /// 空字串還是完全沒有這個 key，都會得到一個對應的資料字典 key（找不到節點就給空字串）。
    /// </summary>
    private static Dictionary<string, object> BuildDataDictionary(MappingConfig mapping, TemplateScanResult scanResult, JToken rootToken)
    {
        var data = new Dictionary<string, object> { [ReservedColumnName] = "" };

        foreach (var fieldKey in scanResult.Fields)
        {
            mapping.Fields.TryGetValue(fieldKey, out var configuredPath);
            var token = rootToken.SelectToken(EffectivePath(fieldKey, configuredPath));
            data[fieldKey] = token == null ? "" : TokenToValue(token);
        }

        foreach (var collectionScan in scanResult.Collections)
        {
            var collectionMapping = mapping.Collections.FirstOrDefault(c => c.Name == collectionScan.Name);
            var collectionPath = EffectivePath(collectionScan.Name, collectionMapping?.Path);
            var arrayToken = rootToken.SelectToken(collectionPath);
            var rows = new List<IDictionary<string, object>>();

            if (arrayToken != null && arrayToken.Type == JTokenType.Array)
            {
                var spread = collectionMapping?.Spread;
                foreach (var item in arrayToken)
                {
                    var row = new Dictionary<string, object> { [ReservedColumnName] = "" };
                    foreach (var columnKey in collectionScan.Columns)
                    {
                        string? configuredColumnPath = null;
                        collectionMapping?.Columns.TryGetValue(columnKey, out configuredColumnPath);
                        var valueToken = item.SelectToken(EffectivePath(columnKey, configuredColumnPath));
                        row[columnKey] = valueToken == null ? "" : TokenToValue(valueToken);
                    }

                    if (spread != null)
                    {
                        ApplySpread(item, spread, row);
                    }

                    rows.Add(row);
                }
            }
            // 找不到對應節點（或節點不是陣列）時，視為沒有明細列，不中斷報表產生。

            data[collectionScan.Name] = rows;
        }

        return data;
    }

    private static void ApplySpread(JToken item, SpreadConfig spread, Dictionary<string, object> row)
    {
        // SelectTokens（複數）同時涵蓋兩種來源寫法：
        // - "$.values" 這種指向單一陣列節點的路徑，會回傳「一個」token（該陣列本身），此時展開其元素；
        // - "$.samples[*].measuredValue" 這種帶萬用字元、跨多個節點取值的路徑，會直接回傳「多個」純量 token。
        var tokens = item.SelectTokens(spread.From).ToList();
        var values = tokens.Count == 1 && tokens[0].Type == JTokenType.Array
            ? tokens[0].Select(TokenToValue).ToList()
            : tokens.Select(TokenToValue).ToList();

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

        var imagePart = drawingsPart.AddImagePart(DetectImagePartType(imageBytes));
        using (var imageStream = new MemoryStream(imageBytes))
        {
            imagePart.FeedData(imageStream);
        }
        var imageRelId = drawingsPart.GetIdOfPart(imagePart);

        var (col, row) = ParseCellReference(cellReference);
        var picId = (uint)(worksheetDrawing.Elements<Xdr.OneCellAnchor>().Count() + 1);

        var anchor = new Xdr.OneCellAnchor(
            new Xdr.FromMarker
            {
                ColumnId = new Xdr.ColumnId(col.ToString()),
                ColumnOffset = new Xdr.ColumnOffset("0"),
                RowId = new Xdr.RowId(row.ToString()),
                RowOffset = new Xdr.RowOffset("0")
            },
            new Xdr.Extent { Cx = ImageWidthEmu, Cy = ImageHeightEmu },
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
                        new A.Extents { Cx = ImageWidthEmu, Cy = ImageHeightEmu }),
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

    private async Task<byte[]> LoadTemplateBytesAsync(ReportConfig config)
    {
        var templatePath = Path.Combine(_env.ContentRootPath, "App_Data", "templates", config.TemplateFile);
        if (!File.Exists(templatePath))
        {
            throw new ReportGenerationException($"找不到範本檔案：{config.TemplateFile}");
        }

        var lastWriteTicks = File.GetLastWriteTimeUtc(templatePath).Ticks;
        var cacheKey = $"tpl:{config.TemplateFile}:{lastWriteTicks}";
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
