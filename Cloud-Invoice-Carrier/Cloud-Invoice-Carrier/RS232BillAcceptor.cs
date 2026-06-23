using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Ports;
using System.Linq;

namespace Cloud_Invoice_Carrier;

/// <summary>
/// RS232-ICT004 紙鈔機串口監聽（參考 puddingDog01 PhotoBoothWin）。
/// </summary>
public sealed class RS232BillAcceptor : IDisposable
{
    private const int BillAcceptCooldownMs = 800;

    private SerialPort? _serialPort;
    private bool _disposed;
    private readonly string _portName;
    private readonly int _baudRate;
    private readonly Queue<byte> _receiveBuffer = new();
    private ReceiveState _currentState = ReceiveState.Idle;
    private bool _isEscrowMode;
    private DateTime _escrowStartTime;
    private DateTime _lastBillAcceptedUtc = DateTime.MinValue;

    private readonly Dictionary<byte, int> _denominationMap = new()
    {
        { 0x10, 100 },
        { 0x20, 200 },
        { 0x50, 500 },
        { 0x64, 100 },
        { 0xC8, 200 },
        { 0x40, 100 },
        { 0x41, 200 },
        { 0x42, 500 },
        { 0x43, 1000 },
        { 0x44, 2000 },
        { 0x31, 100 },
        { 0x32, 200 },
        { 0x35, 500 },
        { 0x00, 100 },
        { 0x01, 100 },
        { 0x02, 200 },
        { 0x05, 500 },
        { 0x0A, 1000 },
        { 0x11, 100 },
        { 0x12, 200 },
        { 0x15, 500 },
        { 0x1A, 1000 },
        { 0x14, 2000 },
        { 0x24, 2000 },
        { 0xF4, 500 },
        { 0x68, 100 },
        { 0xAE, 100 },
        { 0x6C, 100 }
    };

    public event EventHandler<int>? BillReceived;
    public event EventHandler<string>? StatusChanged;
    public event EventHandler<string>? ErrorOccurred;

    private enum ReceiveState
    {
        Idle,
        WaitingForDenomination
    }

    public bool IsOpen => _serialPort != null && _serialPort.IsOpen;

    public string ConfiguredPortName => _portName;

    public string? ConnectedPortName => _serialPort?.PortName;

    public RS232BillAcceptor(string portName = "COM3", int baudRate = 9600)
    {
        _portName = portName;
        _baudRate = baudRate;
    }

    public void Start()
    {
        if (_serialPort != null && _serialPort.IsOpen)
            return;

        try
        {
            var availablePorts = SerialPort.GetPortNames();
            LogMessage($"可用串口：{(availablePorts.Length == 0 ? "(無)" : string.Join(", ", availablePorts))}");

            if (!TryFindPort(out var foundPort))
            {
                var available = SerialPort.GetPortNames();
                var configured = _portName?.Trim();
                var message = !string.IsNullOrWhiteSpace(configured)
                    ? $"找不到設定的 {configured}。目前可用：{(available.Length == 0 ? "(無)" : string.Join(", ", available))}"
                    : "找不到紙鈔機可用串口（已排除 COM8 投幣器保留埠）";
                LogMessage(message);
                ErrorOccurred?.Invoke(this, message);
                return;
            }

            if (IsPortInUse(foundPort))
                LogMessage($"警告：串口 {foundPort} 可能被其他程式占用（例如 RS232-ICT004.exe）");

            _serialPort = new SerialPort(foundPort, _baudRate, Parity.Even, 8, StopBits.One)
            {
                Handshake = Handshake.None,
                ReadTimeout = 1000,
                WriteTimeout = 3000,
                DtrEnable = true,
                RtsEnable = true
            };

            _serialPort.DataReceived += SerialPort_DataReceived;
            _serialPort.ErrorReceived += SerialPort_ErrorReceived;
            _serialPort.Open();

            LogMessage($"串口已開啟：{foundPort}，{_baudRate} bps，Even，DTR/RTS=ON");
            Thread.Sleep(200);

            StatusChanged?.Invoke(this, $"紙鈔機串口就緒（{foundPort}）");
            LogMessage("串口已開啟，等待前端下達啟用指令（待機頁）…");
        }
        catch (UnauthorizedAccessException ex)
        {
            var message = $"串口訪問被拒絕：{ex.Message}（請關閉 RS232-ICT004.exe 或其他占用 COM 的程式）";
            LogMessage(message);
            ErrorOccurred?.Invoke(this, message);
        }
        catch (Exception ex)
        {
            var message = $"開啟串口失敗：{ex.Message}";
            LogMessage(message);
            ErrorOccurred?.Invoke(this, message);
        }
    }

    private static bool IsPortInUse(string portName)
    {
        try
        {
            using var testPort = new SerialPort(portName);
            testPort.Open();
            testPort.Close();
            return false;
        }
        catch
        {
            return true;
        }
    }

    private bool TryFindPort(out string portName)
    {
        portName = string.Empty;
        var ports = SerialPort.GetPortNames();
        if (ports.Length == 0)
            return false;

        var configured = _portName?.Trim();
        if (!string.IsNullOrWhiteSpace(configured))
        {
            if (IsExcludedBillAcceptorPort(configured))
                return false;

            var match = ports.FirstOrDefault(p => p.Equals(configured, StringComparison.OrdinalIgnoreCase));
            if (!string.IsNullOrWhiteSpace(match))
            {
                portName = match;
                return true;
            }

            return false;
        }

        var fallback = ports
            .Where(p => !IsExcludedBillAcceptorPort(p))
            .OrderBy(p =>
            {
                if (p.Length > 3 && int.TryParse(p.AsSpan(3), out var num))
                    return num;
                return int.MaxValue;
            })
            .FirstOrDefault();

        if (string.IsNullOrWhiteSpace(fallback))
            return false;

        portName = fallback;
        return true;
    }

    private static bool IsExcludedBillAcceptorPort(string port) =>
        port.Equals("COM8", StringComparison.OrdinalIgnoreCase);

    private void SerialPort_DataReceived(object sender, SerialDataReceivedEventArgs e)
    {
        if (_serialPort == null || !_serialPort.IsOpen)
            return;

        try
        {
            var bytesToRead = _serialPort.BytesToRead;
            if (bytesToRead == 0)
                return;

            var buffer = new byte[bytesToRead];
            var bytesRead = _serialPort.Read(buffer, 0, bytesToRead);
            for (var i = 0; i < bytesRead; i++)
                _receiveBuffer.Enqueue(buffer[i]);

            ProcessBuffer();
        }
        catch (Exception ex)
        {
            LogMessage($"讀取串口失敗：{ex.Message}");
            ErrorOccurred?.Invoke(this, $"讀取串口數據失敗：{ex.Message}");
        }
    }

    private void SerialPort_ErrorReceived(object sender, SerialErrorReceivedEventArgs e)
    {
        LogMessage($"串口錯誤：{e.EventType}");
        ErrorOccurred?.Invoke(this, $"串口錯誤：{e.EventType}");
    }

    private void ProcessBuffer()
    {
        if (_isEscrowMode && (DateTime.Now - _escrowStartTime).TotalSeconds > 5)
        {
            LogMessage("Escrow 超時，紙鈔被自動拒收");
            ErrorOccurred?.Invoke(this, "Escrow 超時，紙鈔被自動拒收");
            _isEscrowMode = false;
            _currentState = ReceiveState.Idle;
        }

        while (_receiveBuffer.Count > 0)
        {
            var currentByte = _receiveBuffer.Dequeue();
            LogMessage($"[接收] 0x{currentByte:X2}");
            ProcessByteICTProtocol(currentByte);
        }
    }

    private void ProcessByteICTProtocol(byte data)
    {
        switch (data)
        {
            case 0x80:
            case 0x8F:
                LogMessage("收到紙鈔機上電訊號");
                StatusChanged?.Invoke(this, "紙鈔機上電，發送回應…");
                SendCommand(new byte[] { 0x02 });
                _currentState = ReceiveState.Idle;
                break;

            case 0x81:
                LogMessage("收到紙鈔驗證成功訊號 [0x81]");
                StatusChanged?.Invoke(this, "紙鈔驗證成功，等待幣值…");
                _isEscrowMode = true;
                _escrowStartTime = DateTime.Now;
                _currentState = ReceiveState.WaitingForDenomination;
                break;

            case 0x10:
            case 0x30:
            case 0x40:
            case 0x41:
            case 0x42:
            case 0x43:
            case 0x44:
            case 0x50:
            case 0x64:
            case 0xC8:
            case 0x31:
            case 0x32:
            case 0x35:
            case 0x68:
            case 0xAE:
            case 0x6C:
                if (_isEscrowMode || _currentState == ReceiveState.WaitingForDenomination)
                {
                    ProcessBillDenomination(data);
                    _isEscrowMode = false;
                    _currentState = ReceiveState.Idle;
                }
                else
                    ProcessBillDenomination(data);
                break;

            case 0x20:
                if (_isEscrowMode || _currentState == ReceiveState.WaitingForDenomination)
                {
                    ProcessBillDenomination(data);
                    _isEscrowMode = false;
                    _currentState = ReceiveState.Idle;
                }
                else
                    StatusChanged?.Invoke(this, "重啟紙鈔機");
                break;

            case 0x21:
                ErrorOccurred?.Invoke(this, "馬達故障");
                break;
            case 0x22:
                ErrorOccurred?.Invoke(this, "校驗和錯誤");
                break;
            case 0x23:
                ErrorOccurred?.Invoke(this, "卡鈔！請檢查紙鈔機");
                break;
            case 0x24:
                StatusChanged?.Invoke(this, "紙鈔被取走");
                break;
            case 0x25:
                ErrorOccurred?.Invoke(this, "錢箱未關閉");
                break;
            case 0x27:
                ErrorOccurred?.Invoke(this, "感應器問題");
                break;
            case 0x28:
                ErrorOccurred?.Invoke(this, "鉤鈔（Bill Fish）");
                break;
            case 0x29:
                ErrorOccurred?.Invoke(this, "錢箱問題");
                break;
            case 0x2A:
                StatusChanged?.Invoke(this, "紙鈔被拒收");
                break;
            case 0x2E:
                ErrorOccurred?.Invoke(this, "無效指令");
                break;
            case 0x2F:
                StatusChanged?.Invoke(this, "保留狀態");
                break;
            case 0x3E:
                StatusChanged?.Invoke(this, "紙鈔機啟用狀態回報");
                break;
            case 0x5E:
                StatusChanged?.Invoke(this, "紙鈔機禁用狀態回報");
                break;
            case 0x71:
            case 0xFC:
            case 0xF8:
            case 0xF6:
            case 0xE0:
                break;

            default:
                if (_isEscrowMode || _currentState == ReceiveState.WaitingForDenomination)
                {
                    if (ParseBillAmount(data) > 0)
                    {
                        ProcessBillDenomination(data);
                        _isEscrowMode = false;
                        _currentState = ReceiveState.Idle;
                    }
                    else
                        LogMessage($"未知字節在 Escrow 模式下：0x{data:X2}");
                }
                else
                    LogMessage($"未知字節：0x{data:X2}");
                break;
        }
    }

    private void ProcessBillDenomination(byte billCode)
    {
        var amount = ParseBillAmount(billCode);
        if (amount <= 0)
        {
            ErrorOccurred?.Invoke(this, $"無法解析的面額代碼: 0x{billCode:X2}");
            return;
        }

        LogMessage($"解析面額：{amount} 元 (0x{billCode:X2})");

        if (amount != 100)
        {
            var reject = $"拒收紙鈔：只接受 100 元，收到 {amount} 元";
            LogMessage(reject);
            StatusChanged?.Invoke(this, reject);
            ErrorOccurred?.Invoke(this, reject);
            return;
        }

        var now = DateTime.UtcNow;
        if ((now - _lastBillAcceptedUtc).TotalMilliseconds < BillAcceptCooldownMs)
            return;

        _lastBillAcceptedUtc = now;
        StatusChanged?.Invoke(this, $"收到紙鈔: {amount}元");
        BillReceived?.Invoke(this, amount);
        SendCommand(new byte[] { 0x02 });
        LogMessage("已發送接受訊號 [0x02]");
    }

    public void EnableValidator()
    {
        SendCommand(new byte[] { 0x3E });
        StatusChanged?.Invoke(this, "紙鈔機已啟用");
        LogMessage("已發送啟用指令 [0x3E]");
    }

    public void DisableValidator()
    {
        SendCommand(new byte[] { 0x5E });
        StatusChanged?.Invoke(this, "紙鈔機已禁用");
        LogMessage("已發送禁用指令 [0x5E]");
    }

    public void ResetValidator()
    {
        SendCommand(new byte[] { 0x30 });
        StatusChanged?.Invoke(this, "紙鈔機重置中…");
        LogMessage("已發送重置指令 [0x30]");
        Thread.Sleep(100);
    }

    private void SendCommand(byte[] command)
    {
        try
        {
            if (_serialPort == null || !_serialPort.IsOpen)
                return;

            _serialPort.Write(command, 0, command.Length);
            _serialPort.BaseStream.Flush();

            var hex = string.Join(" ", Array.ConvertAll(command, b => $"0x{b:X2}"));
            LogMessage($"[發送] {hex}");
        }
        catch (Exception ex)
        {
            LogMessage($"發送指令失敗：{ex.Message}");
            ErrorOccurred?.Invoke(this, $"發送指令失敗：{ex.Message}");
        }
    }

    private int ParseBillAmount(byte billCode)
    {
        if (_denominationMap.TryGetValue(billCode, out var amount))
            return amount;

        if (billCode is >= 0x30 and <= 0x39)
        {
            var digit = billCode - 0x30;
            return digit switch
            {
                1 => 100,
                2 => 200,
                5 => 500,
                0 => 1000,
                _ => digit * 100
            };
        }

        return 0;
    }

    private void LogMessage(string message)
    {
        var logEntry = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] {message}";
        System.Diagnostics.Debug.WriteLine(logEntry);

        try
        {
            var logDir = Path.Combine(AppPaths.ExeDirectory, "Logs");
            Directory.CreateDirectory(logDir);
            var logFile = Path.Combine(logDir, $"BillAcceptor_{DateTime.Now:yyyyMMdd}.log");
            File.AppendAllText(logFile, logEntry + Environment.NewLine);
        }
        catch
        {
            // 記錄失敗不影響主流程
        }
    }

    public void Stop()
    {
        if (_serialPort != null && _serialPort.IsOpen)
        {
            DisableValidator();
            Thread.Sleep(100);
            _serialPort.Close();
            LogMessage("串口已關閉");
            StatusChanged?.Invoke(this, "紙鈔機已停止");
        }

        _currentState = ReceiveState.Idle;
        _isEscrowMode = false;
        _receiveBuffer.Clear();
    }

    public void Dispose()
    {
        if (_disposed)
            return;

        Stop();
        if (_serialPort != null)
        {
            _serialPort.DataReceived -= SerialPort_DataReceived;
            _serialPort.ErrorReceived -= SerialPort_ErrorReceived;
            _serialPort.Dispose();
        }

        _disposed = true;
    }
}
