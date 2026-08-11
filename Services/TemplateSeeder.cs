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
    private const int LabelSpan = 2;
    private const int ValueSpan = 5;
    private const int FieldsPerRow = 3;
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
        worksheet.Append(new SheetDimension { Reference = $"A1:{ColumnLetter(TotalColumns)}10" });
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

        // 重複標題列 $1:$9，讓多頁列印時每頁都帶著表頭與明細表頭。
        var definedNames = new DefinedNames();
        var printTitles = new DefinedName
        {
            Name = "_xlnm.Print_Titles",
            Text = $"'{sheet.Name}'!$1:$9"
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
        row.Append(CreateInlineStringCell("A1", text, 1));
        for (var col = 2; col <= TotalColumns; col++)
        {
            row.Append(CreateInlineStringCell(CellRef(col, 1), "", 1));
        }
        sheetData.Append(row);
        mergeCells.Append(new MergeCell { Reference = $"A1:{CellRef(TotalColumns, 1)}" });
    }

    /// <summary>
    /// 保留標記，永遠對應空字串，附加在每個單值標記後面讓儲存格內含「多個標記」。
    /// MiniExcel 套版時，單一標記且內容像數字（例如 "0040"）會被判斷成數值型別而遺失前導零；
    /// 只要儲存格內有一個以上的標記，MiniExcel 就會保留為文字型別，藉此避免此問題。
    /// </summary>
    private const string StrGuardMarker = "{{_str}}";

    private static string WithStrGuard(string marker) => marker + StrGuardMarker;

    private static readonly (string Label, string Placeholder)[] HeaderFields =
    {
        ("單據號", "{{docNo}}"), ("序號", "{{seqNo}}"), ("來源單據", "{{sourceDoc}}"),
        ("品號", "{{partNo}}"), ("QC模組", "{{qcModule}}"), ("製程代碼", "{{processCode}}"),
        ("數量", "{{qty}}"), ("品名", "{{partName}}"), ("製程名稱", "{{processName}}"),
        ("規格", "{{spec}}"),
        ("檢驗點", "{{inspectPoint}}"), ("廠商", "{{vendor}}"), ("備註", "{{remark}}"),
        ("日期", "{{date}}"), ("檢驗備註", "{{inspectRemark}}"), ("檢驗規範", "{{inspectSpec}}"),
    };

    private static void AppendFieldRows(SheetData sheetData, MergeCells mergeCells)
    {
        // Row2-4: 每列 3 組 label/value；Row5: 規格獨占整列；Row6-7: 每列 3 組；Row8: 留白緩衝列。
        var rowIndex = 2;
        var fieldIdx = 0;

        void AppendTripleRow(int rIdx)
        {
            var row = new Row { RowIndex = (uint)rIdx, Height = 20, CustomHeight = true };
            var col = 1;
            for (var f = 0; f < FieldsPerRow; f++)
            {
                var (label, placeholder) = HeaderFields[fieldIdx++];
                var labelStart = col;
                var labelEnd = col + LabelSpan - 1;
                row.Append(CreateInlineStringCell(CellRef(labelStart, rIdx), label, 2));
                for (var c = labelStart + 1; c <= labelEnd; c++)
                {
                    row.Append(CreateInlineStringCell(CellRef(c, rIdx), "", 2));
                }
                if (labelEnd > labelStart)
                {
                    mergeCells.Append(new MergeCell { Reference = $"{CellRef(labelStart, rIdx)}:{CellRef(labelEnd, rIdx)}" });
                }

                var valueStart = labelEnd + 1;
                var valueEnd = valueStart + ValueSpan - 1;
                row.Append(CreateInlineStringCell(CellRef(valueStart, rIdx), WithStrGuard(placeholder), 3));
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

        void AppendSpecRow(int rIdx)
        {
            var row = new Row { RowIndex = (uint)rIdx, Height = 24, CustomHeight = true };
            var (label, placeholder) = HeaderFields[fieldIdx++];
            row.Append(CreateInlineStringCell(CellRef(1, rIdx), label, 2));
            row.Append(CreateInlineStringCell(CellRef(2, rIdx), "", 2));
            mergeCells.Append(new MergeCell { Reference = $"{CellRef(1, rIdx)}:{CellRef(2, rIdx)}" });

            row.Append(CreateInlineStringCell(CellRef(3, rIdx), WithStrGuard(placeholder), 3));
            for (var c = 4; c <= TotalColumns; c++)
            {
                row.Append(CreateInlineStringCell(CellRef(c, rIdx), "", 3));
            }
            mergeCells.Append(new MergeCell { Reference = $"{CellRef(3, rIdx)}:{CellRef(TotalColumns, rIdx)}" });
            sheetData.Append(row);
        }

        void AppendSpacerRow(int rIdx)
        {
            var row = new Row { RowIndex = (uint)rIdx, Height = 10, CustomHeight = true };
            row.Append(CreateInlineStringCell(CellRef(1, rIdx), "", 3));
            for (var c = 2; c <= TotalColumns; c++)
            {
                row.Append(CreateInlineStringCell(CellRef(c, rIdx), "", 3));
            }
            mergeCells.Append(new MergeCell { Reference = $"{CellRef(1, rIdx)}:{CellRef(TotalColumns, rIdx)}" });
            sheetData.Append(row);
        }

        AppendTripleRow(rowIndex++); // row2
        AppendTripleRow(rowIndex++); // row3
        AppendTripleRow(rowIndex++); // row4
        AppendSpecRow(rowIndex++);   // row5
        AppendTripleRow(rowIndex++); // row6
        AppendTripleRow(rowIndex++); // row7
        AppendSpacerRow(rowIndex++); // row8
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

    private static void AppendDetailHeaderRow(SheetData sheetData)
    {
        var row = new Row { RowIndex = 9, Height = 20, CustomHeight = true };
        for (var c = 1; c <= TotalColumns; c++)
        {
            row.Append(CreateInlineStringCell(CellRef(c, 9), DetailHeaderLabels[c - 1], 4));
        }
        sheetData.Append(row);
    }

    private static void AppendDetailMarkerRow(SheetData sheetData)
    {
        var row = new Row { RowIndex = 10, Height = 18, CustomHeight = true };
        var markers = new List<string>
        {
            "{{items.seq}}", "{{items.itemName}}", "{{items.point}}", "{{items.std}}",
            "{{items.lower}}", "{{items.upper}}", "{{items.gauge}}"
        };
        for (var i = 1; i <= 13; i++)
        {
            markers.Add($"{{{{items.v{i}}}}}");
        }
        markers.Add("{{items.result}}");

        for (var c = 1; c <= TotalColumns; c++)
        {
            row.Append(CreateInlineStringCell(CellRef(c, 10), markers[c - 1] + "{{items._str}}", 5));
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
            new Fill(new PatternFill(new ForegroundColor { Rgb = "FFD9D9D9" }) { PatternType = PatternValues.Solid }) // 2 淺灰
        )
        { Count = 3 };

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
            }
        )
        { Count = 6 };

        return new Stylesheet(fonts, fills, borders, cellFormats);
    }
}
