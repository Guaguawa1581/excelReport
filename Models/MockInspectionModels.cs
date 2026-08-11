namespace excelReport.Models;

/// <summary>Mock 資料來源 API 的回傳外層結構。</summary>
public class MockApiResponse
{
    public bool Success { get; set; }
    public MockInspectionData? Data { get; set; }
}

/// <summary>外部檢驗紀錄表的表頭資料。</summary>
public class MockInspectionData
{
    public string DocNo { get; set; } = "";
    public string SeqNo { get; set; } = "";
    public string SourceDoc { get; set; } = "";
    public string PartNo { get; set; } = "";
    public string QcModule { get; set; } = "";
    public string ProcessCode { get; set; } = "";
    public string Qty { get; set; } = "";
    public string PartName { get; set; } = "";
    public string ProcessName { get; set; } = "";
    public string Spec { get; set; } = "";
    public string InspectPoint { get; set; } = "";
    public string Vendor { get; set; } = "";
    public string Remark { get; set; } = "";
    public string Date { get; set; } = "";
    public string InspectRemark { get; set; } = "";
    public string InspectSpec { get; set; } = "";
    public List<MockInspectionDetail> Details { get; set; } = new();
}

/// <summary>外部檢驗紀錄表的明細列。</summary>
public class MockInspectionDetail
{
    public int Seq { get; set; }
    public string ItemName { get; set; } = "";
    public string Point { get; set; } = "";
    public string Std { get; set; } = "";
    public string Lower { get; set; } = "";
    public string Upper { get; set; } = "";
    public string Gauge { get; set; } = "";
    public List<string> Values { get; set; } = new();
    public string Result { get; set; } = "";
}
