using excelReport.Models;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using Newtonsoft.Json.Serialization;

namespace excelReport.Controllers;

/// <summary>模擬外部檢驗資料來源，僅供本 prototype 自我測試使用。</summary>
[ApiController]
[Route("api/mock")]
public class MockDataController : ControllerBase
{
    private static readonly JsonSerializerSettings JsonSettings = new()
    {
        ContractResolver = new CamelCasePropertyNamesContractResolver(),
        Formatting = Formatting.Indented
    };

    private static readonly string[] ItemNames =
    {
        "外觀檢驗", "尺寸檢驗", "硬度檢驗", "表面處理檢驗", "鍍層厚度檢驗", "刻印檢驗"
    };

    private static readonly string[] Gauges = { "目視", "游標卡尺", "硬度計", "厚度計", "投影機" };

    [HttpGet("inspection/{docNo}")]
    public IActionResult GetInspection(string docNo)
    {
        if (string.IsNullOrWhiteSpace(docNo))
        {
            var badBody = JsonConvert.SerializeObject(new { success = false, message = "docNo 不可為空" }, JsonSettings);
            return Content(badBody, "application/json");
        }

        var data = BuildMockData(docNo);
        var response = new MockApiResponse { Success = true, Data = data };
        var json = JsonConvert.SerializeObject(response, JsonSettings);
        return Content(json, "application/json");
    }

    /// <summary>穩定雜湊（FNV-1a），確保同一 docNo 在不同次啟動下產生相同的模擬資料筆數與內容。
    /// 不可用 string.GetHashCode()，其雜湊值每次程序啟動都會隨機化。</summary>
    private static int StableHash(string text)
    {
        unchecked
        {
            var hash = 2166136261;
            foreach (var c in text)
            {
                hash = (hash ^ c) * 16777619;
            }
            return (int)(hash & 0x7FFFFFFF);
        }
    }

    private static MockInspectionData BuildMockData(string docNo)
    {
        // 依 docNo 產生穩定但可變化的明細筆數與量測值，方便測試版面在不同筆數下的表現。
        var seed = StableHash(docNo);
        var rng = new Random(seed);
        var detailCount = 3 + (seed % 13); // 3 ~ 15 筆

        var details = new List<MockInspectionDetail>();
        for (var i = 1; i <= detailCount; i++)
        {
            var itemName = ItemNames[(seed + i) % ItemNames.Length];
            var gauge = Gauges[(seed + i) % Gauges.Length];
            var valueCount = 2 + (i % 3); // 2 ~ 4 個量測值
            var values = new List<string>();
            var isFail = false;
            for (var v = 0; v < valueCount; v++)
            {
                var ok = rng.NextDouble() > 0.15;
                if (!ok) isFail = true;
                values.Add(ok ? "V" : "X");
            }

            details.Add(new MockInspectionDetail
            {
                Seq = i,
                ItemName = itemName,
                Point = i.ToString(),
                Std = "0",
                Lower = "0",
                Upper = "0",
                Gauge = gauge,
                Values = values,
                Result = isFail ? "不合格" : "合格"
            });
        }

        return new MockInspectionData
        {
            DocNo = docNo,
            SeqNo = "0040",
            SourceDoc = "",
            PartNo = "XN271427AC3",
            QcModule = "IPQC",
            ProcessCode = "H300",
            Qty = "27 PCS",
            PartName = "M0B1-2015010249",
            ProcessName = "BV30熱處理",
            Spec = "十字花刻Arca Xion+Pat ,電亮HRC45-47",
            InspectPoint = "20",
            Vendor = "湘北",
            Remark = "AA_TEST001",
            Date = DateTime.Now.ToString("yyyy/M/d"),
            InspectRemark = $"TEST001_{docNo}",
            InspectSpec = "IPQC-STD",
            Details = details
        };
    }
}
