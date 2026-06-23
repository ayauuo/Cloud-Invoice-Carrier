using System.Diagnostics;
using System.Globalization;
using System.Text.Json;

namespace Cloud_Invoice_Carrier;

public partial class Form1
{
    private void PostHostRpcResponse(string id, bool ok, object? data = null, string? error = null)
    {
        if (webView21.CoreWebView2 == null)
            return;

        var payload = ok
            ? JsonSerializer.Serialize(new { id, ok = true, data })
            : JsonSerializer.Serialize(new { id, ok = false, error });

        webView21.CoreWebView2.PostWebMessageAsString(payload);
    }

    private bool TryHandleHostRpc(string rawJson)
    {
        if (string.IsNullOrWhiteSpace(rawJson))
            return false;

        var json = rawJson;
        if (json.StartsWith('"') && json.EndsWith('"'))
        {
            try
            {
                json = JsonSerializer.Deserialize<string>(json) ?? json;
            }
            catch
            {
                // 維持原字串
            }
        }

        try
        {
            using var doc = JsonDocument.Parse(json);
            var root = doc.RootElement;
            if (!root.TryGetProperty("cmd", out var cmdProp))
                return false;

            var cmd = cmdProp.GetString();
            if (string.IsNullOrWhiteSpace(cmd))
                return false;

            var id = root.TryGetProperty("id", out var idProp) ? idProp.GetString() ?? "" : "";
            var data = root.TryGetProperty("data", out var dataProp) ? dataProp : default;

            switch (cmd)
            {
                case "get_print_records":
                    HandleGetPrintRecords(id, data);
                    return true;
                case "shutdown":
                    HandleShutdown(id);
                    return true;
                case "exit_app":
                    HandleExitApp(id);
                    return true;
                case "log_payment":
                    HandleLogPayment(id, data);
                    return true;
                case "log_print_record":
                    HandleLogPrintRecord(id, data);
                    return true;
                default:
                    return false;
            }
        }
        catch
        {
            return false;
        }
    }

    private void HandleGetPrintRecords(string id, JsonElement data)
    {
        var dateStr = data.ValueKind == JsonValueKind.Object && data.TryGetProperty("date", out var ds)
            ? ds.GetString() ?? ""
            : "";

        var rangeType = data.ValueKind == JsonValueKind.Object && data.TryGetProperty("rangeType", out var rt)
            ? rt.GetString() ?? "day"
            : "day";

        if (string.IsNullOrWhiteSpace(dateStr)
            || !DateTime.TryParse(dateStr, CultureInfo.InvariantCulture, DateTimeStyles.None, out var parseDate))
        {
            parseDate = DateTime.Today;
        }

        if (rangeType != "week" && rangeType != "month")
            rangeType = "day";

        var (rows, totalPrintSheets, totalTestSheets) = CarrierRecordStore.GetPrintRecordsForView(parseDate, rangeType);
        PostHostRpcResponse(id, true, new { rows, totalPrintSheets, totalTestSheets });
    }

    private void HandleShutdown(string id)
    {
        PostHostRpcResponse(id, true, new { });
        try
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = "shutdown",
                Arguments = "/s /t 0",
                UseShellExecute = false,
                CreateNoWindow = true
            });
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"關機失敗：{ex.Message}");
        }
    }

    private void HandleExitApp(string id)
    {
        if (AppEnvConfig.KioskEnabled)
        {
            PostHostRpcResponse(id, false, error: "Kiosk 模式下無法關閉程式");
            return;
        }

        PostHostRpcResponse(id, true, new { });
        BeginInvoke(Close);
    }

    private void HandleLogPayment(string id, JsonElement data)
    {
        var amount = ReadInt(data, "amount", 100);
        var isTest = ReadBool(data, "isTest");
        var err = CarrierRecordStore.InsertRecord(
            "收款",
            DateTime.Now,
            amount,
            AppEnvConfig.CarrierProjectName,
            Environment.MachineName,
            copies: 0,
            isTest: isTest);

        PostHostRpcResponse(id, err == null, err == null ? new { } : null, err);
    }

    private void HandleLogPrintRecord(string id, JsonElement data)
    {
        var templateName = ReadString(data, "templateName", "載具");
        var carrierCode = ReadString(data, "carrierCode");
        if (!string.IsNullOrWhiteSpace(carrierCode))
        {
            templateName = string.IsNullOrWhiteSpace(templateName)
                ? carrierCode
                : $"{templateName} {carrierCode}";
        }

        var amount = ReadInt(data, "amount", 100);
        var copies = ReadInt(data, "copies", 1);
        var isTest = ReadBool(data, "isTest");

        var err = CarrierRecordStore.InsertRecord(
            templateName,
            DateTime.Now,
            amount,
            AppEnvConfig.CarrierProjectName,
            Environment.MachineName,
            copies,
            isTest);

        PostHostRpcResponse(id, err == null, err == null ? new { } : null, err);
    }

    private static string ReadString(JsonElement data, string name, string fallback = "")
    {
        if (data.ValueKind != JsonValueKind.Object || !data.TryGetProperty(name, out var prop))
            return fallback;

        return prop.GetString() ?? fallback;
    }

    private static int ReadInt(JsonElement data, string name, int fallback = 0)
    {
        if (data.ValueKind != JsonValueKind.Object || !data.TryGetProperty(name, out var prop))
            return fallback;

        return prop.ValueKind switch
        {
            JsonValueKind.Number when prop.TryGetInt32(out var n) => n,
            JsonValueKind.String when int.TryParse(prop.GetString(), out var parsed) => parsed,
            _ => fallback
        };
    }

    private static bool ReadBool(JsonElement data, string name)
    {
        if (data.ValueKind != JsonValueKind.Object || !data.TryGetProperty(name, out var prop))
            return false;

        return prop.ValueKind switch
        {
            JsonValueKind.True => true,
            JsonValueKind.False => false,
            JsonValueKind.Number => prop.GetInt32() != 0,
            JsonValueKind.String => prop.GetString() is "1" or "true" or "yes" or "on",
            _ => false
        };
    }

    private void LogCarrierPayment(int amount, bool isTest = false)
    {
        CarrierRecordStore.InsertRecord(
            "收款",
            DateTime.Now,
            amount,
            AppEnvConfig.CarrierProjectName,
            Environment.MachineName,
            copies: 0,
            isTest: isTest);
    }
}
