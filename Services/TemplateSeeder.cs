using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace excelReport.Services;

/// <summary>
/// 應用程式啟動時，若 App_Data/templates 目錄是空的，自動用 OpenXML 產生一份示範範本
/// sample_inspection.xlsx（含表頭標記、明細標記列、框線、欄寬與列印版面設定）。
/// 之所以直接用 OpenXML 而非 MiniExcel 產生範本本身，是因為範本需要精確控制合併儲存格、
/// 框線與列印設定（重複標題列/縮放頁寬），MiniExcel 的簡易寫入 API 不足以表達這些版面需求。
/// </summary>
public static class TemplateSeeder
{
    private const int TotalColumns = 21; // 7 個明細固定欄 + 13 個量測欄 + 1 個結果欄

    public static void EnsureSampleTemplate(string templatesDir, string fileName)
    {
        Directory.CreateDirectory(templatesDir);

        var hasAnyTemplate = Directory.EnumerateFiles(templatesDir, "*.xlsx").Any();
        if (hasAnyTemplate)
        {
            return;
        }

        var path = Path.Combine(templatesDir, fileName);
        BuildTemplate(path);
    }

    private static void BuildTemplate(string path)
    {
        using var document = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook);
        var workbookPart = document.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();

        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        var sheetData = new SheetData();

        var stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
        stylesPart.Stylesheet = BuildStylesheet();
        stylesPart.Stylesheet.Save();

        var mergeCells = new MergeCells();
        var columns = BuildColumns();

        AppendCompanyHeaderRow(sheetData, mergeCells);
        AppendFieldRows(sheetData, mergeCells);
        AppendDetailHeaderRow(sheetData);
        AppendDetailMarkerRow(sheetData);

        var worksheet = new Worksheet();
        // 沒有這個設定，Excel 會忽略 pageSetup 的 fitToWidth/fitToHeight，改用 scale=100%，
        // 導致友列印時被硬切成好幾頁寬。
        worksheet.Append(new SheetProperties { PageSetupProperties = new PageSetupProperties { FitToPage = true } });
        worksheet.Append(new SheetDimension { Reference = $"A1:{ColumnLetter(TotalColumns)}{DetailMarkerRowIndex}" });
        worksheet.Append(new SheetFormatProperties { DefaultRowHeight = 18, DefaultColumnWidth = 10 });
        worksheet.Append(columns);
        worksheet.Append(sheetData);
        if (mergeCells.Any())
        {
            worksheet.Append(mergeCells);
        }

        worksheet.Append(new PageMargins
        {
            Left = 0.3, Right = 0.3, Top = 0.5, Bottom = 0.5, Header = 0.2, Footer = 0.2
        });
        worksheet.Append(BuildPageSetup());
        worksheet.Append(BuildHeaderFooter());

        worksheetPart.Worksheet = worksheet;

        var sheets = workbookPart.Workbook.AppendChild(new Sheets());
        var sheet = new Sheet
        {
            Id = workbookPart.GetIdOfPart(worksheetPart),
            SheetId = 1,
            Name = "外部檢驗紀錄表"
        };
        sheets.Append(sheet);

        // 重複標題列 $1:$10，讓多頁列印時每頁都帶著表頭與明細表頭。
        var definedNames = new DefinedNames();
        var printTitles = new DefinedName
        {
            Name = "_xlnm.Print_Titles",
            Text = $"'{sheet.Name}'!$1:${DetailHeaderRowIndex}"
        };
        definedNames.Append(printTitles);
        workbookPart.Workbook.Append(definedNames);

        workbookPart.Workbook.Save();
    }

    private static Columns BuildColumns()
    {
        var columns = new Columns();
        double[] widths =
        {
            6,  // 項
            18, // 檢驗項目
            8,  // 點位
            8,  // 標準值
            8,  // 下限
            8,  // 上限
            10, // 量具
        };

        for (uint i = 1; i <= widths.Length; i++)
        {
            columns.Append(new Column { Min = i, Max = i, Width = widths[i - 1], CustomWidth = true });
        }

        // v1 ~ v13
        columns.Append(new Column { Min = 8, Max = 20, Width = 6, CustomWidth = true });
        // 檢驗結果
        columns.Append(new Column { Min = 21, Max = 21, Width = 10, CustomWidth = true });

        return columns;
    }

    private static void AppendCompanyHeaderRow(SheetData sheetData, MergeCells mergeCells)
    {
        var row = new Row { RowIndex = 1, Height = 60, CustomHeight = true };
        var text = "銳泰精密工具股份有限公司\nRE-DAI PRECISION TOOLS CO., LTD\n外部檢驗紀錄表";

        // 右上角保留 4 欄給 [[photo]] 圖片標記（方括號語法，MiniExcel 不認得，套版後
        // 由 ReportEngine 另外掃描、下載/解碼並用 OpenXML 嵌入圖片，詳見 ReportEngine.EmbedImagesAsync）。
        const int photoColumns = 4;
        var companyEndCol = TotalColumns - photoColumns;

        row.Append(CreateInlineStringCell("A1", text, 1));
        for (var col = 2; col <= companyEndCol; col++)
        {
            row.Append(CreateInlineStringCell(CellRef(col, 1), "", 1));
        }
        mergeCells.Append(new MergeCell { Reference = $"A1:{CellRef(companyEndCol, 1)}" });

        var photoStart = companyEndCol + 1;
        row.Append(CreateInlineStringCell(CellRef(photoStart, 1), "[[photo]]", 7));
        for (var col = photoStart + 1; col <= TotalColumns; col++)
        {
            row.Append(CreateInlineStringCell(CellRef(col, 1), "", 7));
        }
        mergeCells.Append(new MergeCell { Reference = $"{CellRef(photoStart, 1)}:{CellRef(TotalColumns, 1)}" });

        sheetData.Append(row);
    }

    private static readonly (string Label, string Name)[] HeaderFields =
    {
        ("單據號", "docNo"), ("序號", "seqNo"), ("來源單據", "sourceDoc"),
        ("品號", "partNo"), ("QC模組", "qcModule"), ("製程代碼", "processCode"), ("數量", "qty"),
        ("品名", "partName"), ("製程名稱", "processName"),
        ("規格", "spec"),
        ("檢驗點", "inspectPoint"), ("廠商", "vendor"), ("備註", "remark"), ("日期", "date"),
        ("檢驗備註", "inspectRemark"),
        ("檢驗規範", "inspectSpec"),
    };

    private static readonly Dictionary<string, string> FieldLabelByName =
        HeaderFields.ToDictionary(f => f.Name, f => f.Label);

    private static string CellTextFor(string fieldName) => $"{{{{{fieldName}}}}}";

    /// <summary>
    /// 每列可容納不同數量的 label/value 組，依組數平分欄寬（餘數分給最後一組），
    /// 讓表頭區塊可以是 3 組一列、4 組一列或單組獨占整列（labelSpan 通常較窄）。
    /// </summary>
    private static void AppendFieldRows(SheetData sheetData, MergeCells mergeCells)
    {
        var rowIndex = 2;

        void AppendGroupRow(int rIdx, int height, string[] names, int labelSpan)
        {
            var row = new Row { RowIndex = (uint)rIdx, Height = height, CustomHeight = true };
            var n = names.Length;
            var valueColsTotal = TotalColumns - labelSpan * n;
            var baseValueSpan = valueColsTotal / n;
            var extra = valueColsTotal % n; // 除不盡的欄數分給最後一組
            var col = 1;
            for (var i = 0; i < n; i++)
            {
                var name = names[i];
                var label = FieldLabelByName[name];
                var labelStart = col;
                var labelEnd = col + labelSpan - 1;
                row.Append(CreateInlineStringCell(CellRef(labelStart, rIdx), label, 2));
                for (var c = labelStart + 1; c <= labelEnd; c++)
                {
                    row.Append(CreateInlineStringCell(CellRef(c, rIdx), "", 2));
                }
                if (labelEnd > labelStart)
                {
                    mergeCells.Append(new MergeCell { Reference = $"{CellRef(labelStart, rIdx)}:{CellRef(labelEnd, rIdx)}" });
                }

                var valueSpan = baseValueSpan + (i == n - 1 ? extra : 0);
                var valueStart = labelEnd + 1;
                var valueEnd = valueStart + valueSpan - 1;
                row.Append(CreateInlineStringCell(CellRef(valueStart, rIdx), CellTextFor(name), 3));
                for (var c = valueStart + 1; c <= valueEnd; c++)
                {
                    row.Append(CreateInlineStringCell(CellRef(c, rIdx), "", 3));
                }
                if (valueEnd > valueStart)
                {
                    mergeCells.Append(new MergeCell { Reference = $"{CellRef(valueStart, rIdx)}:{CellRef(valueEnd, rIdx)}" });
                }

                col = valueEnd + 1;
            }
            sheetData.Append(row);
        }

        AppendGroupRow(rowIndex++, 20, new[] { "docNo", "seqNo", "sourceDoc" }, 2);            // row2
        AppendGroupRow(rowIndex++, 20, new[] { "partNo", "qcModule", "processCode", "qty" }, 1); // row3
        AppendGroupRow(rowIndex++, 20, new[] { "partName", "processName" }, 2);                // row4
        AppendGroupRow(rowIndex++, 24, new[] { "spec" }, 2);                                   // row5
        AppendGroupRow(rowIndex++, 20, new[] { "inspectPoint", "vendor", "remark", "date" }, 1); // row6
        AppendGroupRow(rowIndex++, 20, new[] { "inspectRemark" }, 2);                          // row7
        AppendGroupRow(rowIndex++, 20, new[] { "inspectSpec" }, 2);                            // row8

        AppendSectionTitleRow(sheetData, mergeCells, rowIndex++, "子報表");                     // row9
    }

    private static void AppendSectionTitleRow(SheetData sheetData, MergeCells mergeCells, int rIdx, string title)
    {
        var row = new Row { RowIndex = (uint)rIdx, Height = 22, CustomHeight = true };
        row.Append(CreateInlineStringCell(CellRef(1, rIdx), title, 6));
        for (var c = 2; c <= TotalColumns; c++)
        {
            row.Append(CreateInlineStringCell(CellRef(c, rIdx), "", 6));
        }
        mergeCells.Append(new MergeCell { Reference = $"{CellRef(1, rIdx)}:{CellRef(TotalColumns, rIdx)}" });
        sheetData.Append(row);
    }

    private static readonly string[] DetailHeaderLabels = BuildDetailHeaderLabels();

    private static string[] BuildDetailHeaderLabels()
    {
        var labels = new List<string> { "項次", "檢驗項目", "點位", "標準值", "下限", "上限", "量具" };
        for (var i = 1; i <= 13; i++)
        {
            labels.Add(i.ToString());
        }
        labels.Add("檢驗結果");
        return labels.ToArray();
    }

    private const int DetailHeaderRowIndex = 10;
    private const int DetailMarkerRowIndex = 11;

    private static void AppendDetailHeaderRow(SheetData sheetData)
    {
        var row = new Row { RowIndex = DetailHeaderRowIndex, Height = 20, CustomHeight = true };
        for (var c = 1; c <= TotalColumns; c++)
        {
            row.Append(CreateInlineStringCell(CellRef(c, DetailHeaderRowIndex), DetailHeaderLabels[c - 1], 4));
        }
        sheetData.Append(row);
    }

    private static void AppendDetailMarkerRow(SheetData sheetData)
    {
        var row = new Row { RowIndex = DetailMarkerRowIndex, Height = 18, CustomHeight = true };
        var columnNames = new List<string> { "seq", "itemName", "point", "std", "lower", "upper", "gauge" };
        for (var i = 1; i <= 13; i++)
        {
            columnNames.Add($"v{i}");
        }
        columnNames.Add("result");

        for (var c = 1; c <= TotalColumns; c++)
        {
            row.Append(CreateInlineStringCell(CellRef(c, DetailMarkerRowIndex), $"{{{{items.{columnNames[c - 1]}}}}}", 5));
        }
        sheetData.Append(row);
    }

    private static PageSetup BuildPageSetup()
    {
        return new PageSetup
        {
            Orientation = OrientationValues.Landscape,
            FitToWidth = 1,
            FitToHeight = 0,
            PaperSize = 9 // A4
        };
    }

    private static HeaderFooter BuildHeaderFooter()
    {
        return new HeaderFooter
        {
            OddFooter = new OddFooter { Text = "&C第 &P 頁，共 &N 頁" }
        };
    }

    private static Cell CreateInlineStringCell(string reference, string text, uint styleIndex)
    {
        // MiniExcel 的範本掃描是找 t="str" 搭配 <v> 存放文字的儲存格（而非 inlineStr/<is>），
        // 因此範本標記文字一律用這種格式寫入，才能被 SaveAsByTemplate 正確辨識與取代。
        return new Cell
        {
            CellReference = reference,
            DataType = CellValues.String,
            StyleIndex = styleIndex,
            CellValue = new CellValue(text)
        };
    }

    private static string CellRef(int col, int row) => $"{ColumnLetter(col)}{row}";

    private static string ColumnLetter(int col)
    {
        var letter = "";
        while (col > 0)
        {
            var rem = (col - 1) % 26;
            letter = (char)('A' + rem) + letter;
            col = (col - 1) / 26;
        }
        return letter;
    }

    private static Stylesheet BuildStylesheet()
    {
        var fonts = new Fonts(
            new Font(new FontSize { Val = 11 }, new FontName { Val = "微軟正黑體" }), // 0 default
            new Font(new Bold(), new FontSize { Val = 14 }, new FontName { Val = "微軟正黑體" }), // 1 title
            new Font(new Bold(), new FontSize { Val = 11 }, new FontName { Val = "微軟正黑體" }), // 2 label/header
            new Font(new FontSize { Val = 10 }, new FontName { Val = "微軟正黑體" }) // 3 detail cell
        )
        { Count = 4 };

        var fills = new Fills(
            new Fill(new PatternFill { PatternType = PatternValues.None }), // 0
            new Fill(new PatternFill { PatternType = PatternValues.Gray125 }), // 1 (必須保留)
            new Fill(new PatternFill(new ForegroundColor { Rgb = "FFD9D9D9" }) { PatternType = PatternValues.Solid }), // 2 淺灰
            new Fill(new PatternFill(new ForegroundColor { Rgb = "FFFFFF00" }) { PatternType = PatternValues.Solid }) // 3 黃色（子報表標題列）
        )
        { Count = 4 };

        var allThinBorder = new Border(
            new LeftBorder(new Color { Auto = true }) { Style = BorderStyleValues.Thin },
            new RightBorder(new Color { Auto = true }) { Style = BorderStyleValues.Thin },
            new TopBorder(new Color { Auto = true }) { Style = BorderStyleValues.Thin },
            new BottomBorder(new Color { Auto = true }) { Style = BorderStyleValues.Thin },
            new DiagonalBorder()
        );
        var borders = new Borders(
            new Border(new LeftBorder(), new RightBorder(), new TopBorder(), new BottomBorder(), new DiagonalBorder()), // 0 無框線
            allThinBorder // 1 全框線
        )
        { Count = 2 };

        var cellFormats = new CellFormats(
            new CellFormat { FontId = 0, FillId = 0, BorderId = 0 }, // 0 default
            new CellFormat // 1 title
            {
                FontId = 1, FillId = 0, BorderId = 1, ApplyFont = true, ApplyBorder = true, ApplyAlignment = true,
                Alignment = new Alignment { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center, WrapText = true }
            },
            new CellFormat // 2 label
            {
                FontId = 2, FillId = 2, BorderId = 1, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true,
                Alignment = new Alignment { Horizontal = HorizontalAlignmentValues.Left, Vertical = VerticalAlignmentValues.Center }
            },
            new CellFormat // 3 value
            {
                FontId = 0, FillId = 0, BorderId = 1, ApplyBorder = true, ApplyAlignment = true,
                Alignment = new Alignment { Horizontal = HorizontalAlignmentValues.Left, Vertical = VerticalAlignmentValues.Center, WrapText = true }
            },
            new CellFormat // 4 detail header
            {
                FontId = 2, FillId = 2, BorderId = 1, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true,
                Alignment = new Alignment { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center, WrapText = true }
            },
            new CellFormat // 5 detail cell
            {
                FontId = 3, FillId = 0, BorderId = 1, ApplyFont = true, ApplyBorder = true, ApplyAlignment = true,
                Alignment = new Alignment { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center }
            },
            new CellFormat // 6 子報表 section title（黃底）
            {
                FontId = 2, FillId = 3, BorderId = 1, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true,
                Alignment = new Alignment { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center }
            },
            new CellFormat // 7 photo 佔位格
            {
                FontId = 0, FillId = 0, BorderId = 1, ApplyBorder = true, ApplyAlignment = true,
                Alignment = new Alignment { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center, WrapText = true }
            }
        )
        { Count = 8 };

        return new Stylesheet(fonts, fills, borders, cellFormats);
    }
}
