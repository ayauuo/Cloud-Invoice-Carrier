using System.Globalization;

namespace Cloud_Invoice_Carrier;

/// <summary>從程式目錄讀取 .env（KEY=VALUE），用於切換功能與 TSC 參數。</summary>
internal static class AppEnvConfig
{
    public enum AppMode
    {
        Carrier,
        NameLabel
    }

    public static AppMode Mode { get; private set; } = AppMode.Carrier;

    /// <summary>Windows「印表機與掃描器」中顯示的 TSC 印表機名稱。</summary>
    public static string TscWindowsPrinterName { get; private set; } = "TSC TTP-345";

    public static int TscDpi { get; private set; } = 300;

    public static double LabelWidthMm { get; private set; } = 50;

    public static double LabelHeightMm { get; private set; } = 30;

    public static double LabelGapMm { get; private set; } = 2;

    public static void Load(string startupDirectory)
    {
        var path = Path.Combine(startupDirectory, ".env");
        if (!File.Exists(path))
            return;

        foreach (var rawLine in File.ReadAllLines(path))
        {
            var line = rawLine.Trim();
            if (line.Length == 0 || line.StartsWith('#'))
                continue;

            var eq = line.IndexOf('=');
            if (eq <= 0)
                continue;

            var key = line[..eq].Trim();
            var value = line[(eq + 1)..].Trim().Trim('"', '\'');
            if (value.Length == 0)
                continue;

            Apply(key, value);
        }
    }

    private static void Apply(string key, string value)
    {
        switch (key.ToUpperInvariant())
        {
            case "APP_MODE":
                Mode = value.Equals("name_label", StringComparison.OrdinalIgnoreCase)
                    || value.Equals("namelabel", StringComparison.OrdinalIgnoreCase)
                    || value.Equals("tsc", StringComparison.OrdinalIgnoreCase)
                    ? AppMode.NameLabel
                    : AppMode.Carrier;
                break;
            case "TSC_WINDOWS_PRINTER_NAME":
            case "TSC_PRINTER_NAME":
                TscWindowsPrinterName = value;
                break;
            case "TSC_DPI":
                if (int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var dpi) && dpi > 0)
                    TscDpi = dpi;
                break;
            case "LABEL_WIDTH_MM":
                if (double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var w) && w > 0)
                    LabelWidthMm = w;
                break;
            case "LABEL_HEIGHT_MM":
                if (double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var h) && h > 0)
                    LabelHeightMm = h;
                break;
            case "LABEL_GAP_MM":
                if (double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var g) && g >= 0)
                    LabelGapMm = g;
                break;
        }
    }
}
