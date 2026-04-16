using System.Runtime.InteropServices;
using System.Text;

namespace Cloud_Invoice_Carrier;

/// <summary>將 RAW 位元組送到 Windows 印表機佇列（驅動需支援 TSPL 直通）。</summary>
internal static class RawPrinterHelper
{
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private class DocInfo
    {
        [MarshalAs(UnmanagedType.LPWStr)]
        public string DocName = "RAW";

        [MarshalAs(UnmanagedType.LPWStr)]
        public string? OutputFile;

        [MarshalAs(UnmanagedType.LPWStr)]
        public string DataType = "RAW";
    }

    [DllImport("winspool.Drv", EntryPoint = "OpenPrinterW", SetLastError = true, CharSet = CharSet.Unicode)]
    private static extern bool OpenPrinter(string? src, out IntPtr hPrinter, IntPtr pd);

    [DllImport("winspool.Drv", EntryPoint = "ClosePrinter", SetLastError = true)]
    private static extern bool ClosePrinter(IntPtr hPrinter);

    [DllImport("winspool.Drv", EntryPoint = "StartDocPrinterW", SetLastError = true, CharSet = CharSet.Unicode)]
    private static extern bool StartDocPrinter(IntPtr hPrinter, int level, [In] DocInfo di);

    [DllImport("winspool.Drv", EntryPoint = "EndDocPrinter", SetLastError = true)]
    private static extern bool EndDocPrinter(IntPtr hPrinter);

    [DllImport("winspool.Drv", EntryPoint = "StartPagePrinter", SetLastError = true)]
    private static extern bool StartPagePrinter(IntPtr hPrinter);

    [DllImport("winspool.Drv", EntryPoint = "EndPagePrinter", SetLastError = true)]
    private static extern bool EndPagePrinter(IntPtr hPrinter);

    [DllImport("winspool.Drv", EntryPoint = "WritePrinter", SetLastError = true)]
    private static extern bool WritePrinter(IntPtr hPrinter, IntPtr pBytes, int dwCount, out int dwWritten);

    public static void SendBytes(string printerName, byte[] data, string documentName = "TSPL")
    {
        if (!OpenPrinter(printerName, out var hPrinter, IntPtr.Zero))
            throw new IOException("OpenPrinter 失敗：" + printerName + "（請確認「印表機與掃描器」名稱與 .env 的 TSC_WINDOWS_PRINTER_NAME 一致）。");

        try
        {
            var di = new DocInfo { DocName = documentName, DataType = "RAW" };
            if (!StartDocPrinter(hPrinter, 1, di))
                throw new IOException("StartDocPrinter 失敗。");

            try
            {
                if (!StartPagePrinter(hPrinter))
                    throw new IOException("StartPagePrinter 失敗。");

                try
                {
                    var buf = data;
                    var pinned = GCHandle.Alloc(buf, GCHandleType.Pinned);
                    try
                    {
                        var ptr = pinned.AddrOfPinnedObject();
                        if (!WritePrinter(hPrinter, ptr, buf.Length, out var written) || written != buf.Length)
                            throw new IOException("WritePrinter 失敗或長度不符。");
                    }
                    finally
                    {
                        pinned.Free();
                    }
                }
                finally
                {
                    EndPagePrinter(hPrinter);
                }
            }
            finally
            {
                EndDocPrinter(hPrinter);
            }
        }
        finally
        {
            ClosePrinter(hPrinter);
        }
    }

    public static void SendUtf8Text(string printerName, string tsplCommandsUtf8, string documentName = "TSPL")
    {
        SendBytes(printerName, Encoding.UTF8.GetBytes(tsplCommandsUtf8), documentName);
    }
}
