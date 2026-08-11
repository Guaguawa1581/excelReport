using excelReport.Models;

namespace excelReport.Services;

/// <summary>應用程式啟動時，若尚無任何設定檔，自動 seed 一份對應示範範本的完整設定。</summary>
public static class ConfigSeeder
{
    public const string SampleCode = "IPQC_EXT_INSPECT";

    public static void EnsureSampleConfig(IReportConfigStore store, string baseUrl)
    {
        if (store.GetAll().Count > 0)
        {
            return;
        }

        var config = new ReportConfig
        {
            Code = SampleCode,
            Name = "外部檢驗紀錄表",
            TemplateFile = "sample_inspection.xlsx",
            DataSource = new DataSourceConfig
            {
                Url = $"{baseUrl}/api/mock/inspection/{{docNo}}",
                Method = "GET",
                Headers = new Dictionary<string, string>(),
                TimeoutSec = 30
            },
            Parameters = new List<ParameterDef>
            {
                new() { Key = "docNo", Label = "單據號", Type = "string", Required = true }
            },
            Mapping = new MappingConfig
            {
                Root = "$.data",
                Fields = new Dictionary<string, string>
                {
                    ["docNo"] = "$.docNo",
                    ["seqNo"] = "$.seqNo",
                    ["sourceDoc"] = "$.sourceDoc",
                    ["partNo"] = "$.partNo",
                    ["qcModule"] = "$.qcModule",
                    ["processCode"] = "$.processCode",
                    ["qty"] = "$.qty",
                    ["partName"] = "$.partName",
                    ["processName"] = "$.processName",
                    ["spec"] = "$.spec",
                    ["inspectPoint"] = "$.inspectPoint",
                    ["vendor"] = "$.vendor",
                    ["remark"] = "$.remark",
                    ["date"] = "$.date",
                    ["inspectRemark"] = "$.inspectRemark",
                    ["inspectSpec"] = "$.inspectSpec"
                },
                Collections = new List<CollectionMapping>
                {
                    new()
                    {
                        Name = "items",
                        Path = "$.details",
                        Columns = new Dictionary<string, string>
                        {
                            ["seq"] = "$.seq",
                            ["itemName"] = "$.itemName",
                            ["point"] = "$.point",
                            ["std"] = "$.std",
                            ["lower"] = "$.lower",
                            ["upper"] = "$.upper",
                            ["gauge"] = "$.gauge",
                            ["result"] = "$.result"
                        },
                        Spread = new SpreadConfig
                        {
                            From = "$.values",
                            Prefix = "v",
                            Max = 13
                        }
                    }
                }
            }
        };

        store.Save(config);
    }
}
