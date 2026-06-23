using System.Text.Json.Serialization;

namespace Cloud_Invoice_Carrier.Models;

/// <summary>觀看資料庫畫面用：單筆列印／收款紀錄。</summary>
public sealed class PrintRecordViewRow
{
    [JsonPropertyName("id")]
    public long Id { get; set; }

    [JsonPropertyName("date")]
    public string Date { get; set; } = "";

    [JsonPropertyName("time")]
    public string Time { get; set; } = "";

    [JsonPropertyName("isTest")]
    public bool IsTest { get; set; }

    [JsonPropertyName("amount")]
    public int Amount { get; set; }

    [JsonPropertyName("copies")]
    public int Copies { get; set; }

    [JsonPropertyName("projectName")]
    public string ProjectName { get; set; } = "";

    [JsonPropertyName("machineName")]
    public string MachineName { get; set; } = "";

    [JsonPropertyName("templateName")]
    public string TemplateName { get; set; } = "";
}
