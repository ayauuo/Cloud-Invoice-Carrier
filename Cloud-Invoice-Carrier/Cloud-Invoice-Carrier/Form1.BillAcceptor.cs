using System.IO.Ports;
using System.Text.Json;
using Microsoft.Web.WebView2.Core;

namespace Cloud_Invoice_Carrier;

public partial class Form1
{
    private static readonly TimeSpan BillAcceptorHardwareStartDelay = TimeSpan.FromSeconds(5);
    private static readonly TimeSpan BillAcceptorRetryDelay = TimeSpan.FromSeconds(5);
    private const int BillAcceptorMaxRetries = 3;
    private const int BillAcceptorRetryIntervalMs = 2000;
    private const int BillAcceptorPostOpenDelayMs = 500;

    private RS232BillAcceptor? _billAcceptor;
    private bool _billAcceptorConfigEnabled = true;
    private bool _billAcceptorValidatorEnabled;
    private bool _paymentsEnabled = true;
    private int _billAcceptorStartGeneration;
    private readonly object _billAcceptorStartLock = new();

    private void WireBillAcceptorForCarrierMode()
    {
        _billAcceptorConfigEnabled = AppEnvConfig.BillAcceptorEnabled;
        LogBillAcceptorStartup();

        if (AppEnvConfig.Mode != AppEnvConfig.AppMode.Carrier)
            return;

        if (!_billAcceptorConfigEnabled)
            LogBillAcceptor("紙鈔機已在設定中停用（BILL_ACCEPTOR_ENABLED=false）");
    }

    private static void LogBillAcceptorStartup()
    {
        var envPath = Path.Combine(AppPaths.ExeDirectory, ".env");
        var layoutPath = AppPaths.FindFile("carrier-layout.json")
            ?? AppPaths.FindFile("carrier-layout.example.json")
            ?? "(找不到)";
        LogBillAcceptor(
            $"啟動診斷：模式={AppEnvConfig.Mode}，設定檔 BILL_ACCEPTOR_ENABLED={AppEnvConfig.BillAcceptorEnabled}，偏好埠={AppEnvConfig.BillAcceptorPort}，" +
            $".env={(File.Exists(envPath) ? envPath : "不存在")}，layout={layoutPath}，exe目錄={AppPaths.ExeDirectory}");

        try
        {
            var ports = SerialPort.GetPortNames();
            LogBillAcceptor(ports.Length == 0
                ? "系統 COM 埠：(無) — 請確認紙鈔機 USB/RS232 已連接"
                : $"系統 COM 埠：{string.Join(", ", ports.OrderBy(p => p, StringComparer.OrdinalIgnoreCase))}");
        }
        catch (Exception ex)
        {
            LogBillAcceptor($"讀取 COM 埠失敗：{ex.Message}");
        }

        if (AppEnvConfig.Mode != AppEnvConfig.AppMode.Carrier)
            LogBillAcceptor("警告：APP_MODE 不是 carrier，紙鈔機硬體不會啟動");
    }

    private void EnsureBillAcceptorStartedAfterNavigation()
    {
        if (AppEnvConfig.Mode != AppEnvConfig.AppMode.Carrier)
            return;

        ScheduleBillAcceptorHardwareStart();
    }

    private void ScheduleBillAcceptorHardwareStart()
    {
        if (!_billAcceptorConfigEnabled)
            return;

        if (_billAcceptor != null && _billAcceptor.IsOpen)
            return;

        int generation;
        lock (_billAcceptorStartLock)
        {
            generation = ++_billAcceptorStartGeneration;
        }

        LogBillAcceptor($"將於 {BillAcceptorHardwareStartDelay.TotalSeconds:0} 秒後啟動紙鈔機（等待 WebView／USB 穩定）…");

        _ = Task.Run(async () =>
        {
            await Task.Delay(BillAcceptorHardwareStartDelay);

            lock (_billAcceptorStartLock)
            {
                if (generation != _billAcceptorStartGeneration)
                    return;
            }

            StartBillAcceptor();
        });
    }

    private void StartBillAcceptor()
    {
        if (!_billAcceptorConfigEnabled || AppEnvConfig.Mode != AppEnvConfig.AppMode.Carrier)
            return;

        _ = Task.Run(async () =>
        {
            for (var attempt = 1; attempt <= BillAcceptorMaxRetries; attempt++)
            {
                RS232BillAcceptor? acceptor = null;
                string? connectedPort = null;

                try
                {
                    var ports = SerialPort.GetPortNames();
                    LogBillAcceptor(
                        $"紙鈔機初始化 ({attempt}/{BillAcceptorMaxRetries})，COM 埠：{(ports.Length == 0 ? "(無)" : string.Join(", ", ports))}");

                    var portToUse = ResolveBillAcceptorPort();
                    LogBillAcceptor($"嘗試連接紙鈔機：{portToUse}（僅使用 .env 指定埠，不自動改其他 COM）");

                    lock (_billAcceptorStartLock)
                    {
                        DisposeBillAcceptorCore();

                        acceptor = new RS232BillAcceptor(portToUse, 9600);
                        acceptor.BillReceived += OnBillReceived;
                        acceptor.StatusChanged += OnBillAcceptorStatusChanged;
                        acceptor.ErrorOccurred += OnBillAcceptorErrorOccurred;
                        acceptor.Start();
                        _billAcceptor = acceptor;
                    }

                    await Task.Delay(BillAcceptorPostOpenDelayMs);

                    if (acceptor.IsOpen)
                    {
                        connectedPort = acceptor.ConnectedPortName ?? portToUse;
                        LogBillAcceptor($"紙鈔機已連接：{connectedPort}（詳細協議日誌見 Logs\\BillAcceptor_*.log）");
                        BeginInvoke(ApplyBillAcceptorValidatorState);
                        return;
                    }

                    LogBillAcceptor($"串口未能開啟（{attempt}/{BillAcceptorMaxRetries}）");
                }
                catch (Exception ex)
                {
                    LogBillAcceptor($"啟動紙鈔機失敗：{ex.Message}（{attempt}/{BillAcceptorMaxRetries}）");
                }

                if (attempt < BillAcceptorMaxRetries)
                    await Task.Delay(BillAcceptorRetryIntervalMs);
            }

            LogBillAcceptor($"紙鈔機連接失敗，{BillAcceptorRetryDelay.TotalSeconds:0} 秒後再試");
            BeginInvoke(ScheduleBillAcceptorRetry);
        });
    }

    private static string ResolveBillAcceptorPort()
    {
        var configured = AppEnvConfig.BillAcceptorPort?.Trim();
        if (!string.IsNullOrWhiteSpace(configured))
            return configured;

        try
        {
            var ports = SerialPort.GetPortNames();
            return ports
                .Where(p => p.StartsWith("COM", StringComparison.OrdinalIgnoreCase) && !IsExcludedBillAcceptorPort(p))
                .OrderBy(p =>
                {
                    if (p.Length > 3 && int.TryParse(p.AsSpan(3), out var num))
                        return num;
                    return int.MaxValue;
                })
                .FirstOrDefault() ?? "COM4";
        }
        catch
        {
            return "COM4";
        }
    }

    private static bool IsExcludedBillAcceptorPort(string port) =>
        port.Equals("COM8", StringComparison.OrdinalIgnoreCase);

    private void ScheduleBillAcceptorRetry()
    {
        if (!_billAcceptorConfigEnabled || AppEnvConfig.Mode != AppEnvConfig.AppMode.Carrier)
            return;

        if (_billAcceptor != null && _billAcceptor.IsOpen)
            return;

        var generation = _billAcceptorStartGeneration;
        _ = Task.Run(async () =>
        {
            await Task.Delay(BillAcceptorRetryDelay);
            lock (_billAcceptorStartLock)
            {
                if (generation != _billAcceptorStartGeneration)
                    return;
            }

            if (!_billAcceptorConfigEnabled || AppEnvConfig.Mode != AppEnvConfig.AppMode.Carrier)
                return;

            if (_billAcceptor != null && _billAcceptor.IsOpen)
                return;

            LogBillAcceptor("重新嘗試連接紙鈔機…");
            StartBillAcceptor();
        });
    }

    private void StopBillAcceptor()
    {
        lock (_billAcceptorStartLock)
        {
            _billAcceptorStartGeneration++;
            DisposeBillAcceptorCore();
        }
    }

    private void DisposeBillAcceptorCore()
    {
        if (_billAcceptor == null)
            return;

        _billAcceptor.BillReceived -= OnBillReceived;
        _billAcceptor.StatusChanged -= OnBillAcceptorStatusChanged;
        _billAcceptor.ErrorOccurred -= OnBillAcceptorErrorOccurred;
        _billAcceptor.Dispose();
        _billAcceptor = null;
    }

    private void OnBillReceived(object? sender, int amount)
    {
        if (!_paymentsEnabled)
            return;

        BeginInvoke(() =>
        {
            try
            {
                if (webView21.CoreWebView2 == null)
                    return;

                var message = JsonSerializer.Serialize(new Dictionary<string, object>
                {
                    ["@event"] = "paid",
                    ["event"] = "paid",
                    ["amount"] = amount
                });
                webView21.CoreWebView2.PostWebMessageAsString(message);
                LogCarrierPayment(amount);
            }
            catch (Exception ex)
            {
                LogBillAcceptor($"發送付款事件失敗：{ex.Message}");
            }
        });
    }

    private void OnBillAcceptorStatusChanged(object? sender, string status)
    {
        LogBillAcceptor(status);
    }

    private void OnBillAcceptorErrorOccurred(object? sender, string error)
    {
        LogBillAcceptor($"錯誤：{error}");
    }

    private static void LogBillAcceptor(string message)
    {
        var line = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {message}";
        System.Diagnostics.Debug.WriteLine(line);
        try
        {
            var logPath = Path.Combine(AppPaths.ExeDirectory, "bill-acceptor.log");
            File.AppendAllText(logPath, line + Environment.NewLine);
        }
        catch
        {
            // 記錄失敗不影響主流程
        }
    }

    private void HandleBillAcceptorControl(bool enabled)
    {
        _billAcceptorValidatorEnabled = enabled;
        LogBillAcceptor(enabled
            ? "網頁要求啟用驗钞器（僅待機頁會收鈔）"
            : "網頁要求停用驗钞器（非待機頁或功能關閉）");
        ApplyBillAcceptorValidatorState();
    }

    private void ApplyBillAcceptorValidatorState()
    {
        try
        {
            if (_billAcceptor == null || !_billAcceptor.IsOpen)
            {
                if (_billAcceptorValidatorEnabled)
                    LogBillAcceptor($"驗钞器待啟用：串口尚未連上，狀態={_billAcceptorValidatorEnabled}");
                return;
            }

            if (_billAcceptorValidatorEnabled)
                _billAcceptor.EnableValidator();
            else
                _billAcceptor.DisableValidator();
        }
        catch (Exception ex)
        {
            LogBillAcceptor($"控制紙鈔機失敗：{ex.Message}");
        }
    }

    private void SyncBillAcceptorForIdlePage()
    {
        if (!_billAcceptorConfigEnabled || AppEnvConfig.Mode != AppEnvConfig.AppMode.Carrier)
            return;

        _billAcceptorValidatorEnabled = true;
        ApplyBillAcceptorValidatorState();
    }

    private bool TryHandleBillAcceptorWebMessage(string rawJson)
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
            using var jsonDoc = JsonDocument.Parse(json);
            if (!jsonDoc.RootElement.TryGetProperty("@event", out var eventProp))
                return false;

            var eventName = eventProp.GetString();
            if (eventName == "bill_acceptor_control"
                && jsonDoc.RootElement.TryGetProperty("enabled", out var enabledProp))
            {
                HandleBillAcceptorControl(enabledProp.GetBoolean());
                return true;
            }

            if (eventName == "set_payments_enabled"
                && jsonDoc.RootElement.TryGetProperty("enabled", out var paymentsProp))
            {
                _paymentsEnabled = paymentsProp.GetBoolean();
                return true;
            }

            if (eventName == "payments_config")
            {
                if (jsonDoc.RootElement.TryGetProperty("billAcceptorEnabled", out var billProp))
                    _billAcceptorConfigEnabled = billProp.GetBoolean();

                if (_billAcceptorConfigEnabled)
                    ScheduleBillAcceptorHardwareStart();
                else
                    StopBillAcceptor();

                return true;
            }
        }
        catch
        {
            // 非紙鈔機 JSON 訊息
        }

        return false;
    }

    protected override void OnFormClosed(FormClosedEventArgs e)
    {
        StopBillAcceptor();
        base.OnFormClosed(e);
    }
}
