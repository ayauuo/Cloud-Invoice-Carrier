using System.Globalization;

namespace Cloud_Invoice_Carrier;

/// <summary>從程式目錄讀取 .env（KEY=VALUE），用於切換功能與 TSC 參數。</summary>
internal static class AppEnvConfig
{
    public enum NameLabelPrintMode
    {
        Auto,
        Text,
        Bitmap
    }

    public enum AppMode
    {
        Carrier,
        NameLabel
    }

    public enum PaperSensorMode
    {
        Gap,
        BlackMark,
        Continuous
    }

    public static AppMode Mode { get; private set; } = AppMode.Carrier;

    /// <summary>Windows「印表機與掃描器」中顯示的 TSC 印表機名稱。</summary>
    public static string TscWindowsPrinterName { get; private set; } = "TSC TTP-345";

    public static int TscDpi { get; private set; } = 300;

    public static double LabelWidthMm { get; private set; } = 50;

    public static double LabelHeightMm { get; private set; } = 30;

    public static double LabelGapMm { get; private set; } = 2;

    /// <summary>姓名貼排版欄數（橫向）。</summary>
    public static int NameLabelColumns { get; private set; } = 3;

    /// <summary>姓名貼排版列數（縱向）。</summary>
    public static int NameLabelRows { get; private set; } = 6;

    /// <summary>姓名貼格與格間距（毫米）。</summary>
    public static double NameLabelCellGapMm { get; private set; } = 5;

    /// <summary>姓名貼欄與欄間距（毫米）。</summary>
    public static double NameLabelColumnGapMm { get; private set; } = 5;

    /// <summary>姓名貼列與列間距（毫米）。</summary>
    public static double NameLabelRowGapMm { get; private set; } = 5;

    /// <summary>姓名貼排版可用高度（毫米）。0 表示使用整張標籤高度。</summary>
    public static double NameLabelGridHeightMm { get; private set; } = 0;

    /// <summary>姓名貼排版比例（字與間距一起縮放）。1.0 為原始比例。</summary>
    public static double NameLabelLayoutScale { get; private set; } = 1.0;

    /// <summary>姓名貼字元間距（像素）。正值加大、負值縮小。</summary>
    public static int NameLabelCharSpacingPx { get; private set; } = 0;

    /// <summary>第一欄文字水平位移（像素）。正值往右、負值往左。</summary>
    public static int NameLabelFirstColumnOffsetXPx { get; private set; } = 0;

    /// <summary>各欄文字水平位移（像素）。例如 30,-5,10 對應第 1/2/3 欄。</summary>
    public static int[] NameLabelColumnOffsetsXPx { get; private set; } = Array.Empty<int>();

    /// <summary>姓名貼輸出模式：auto / text / bitmap。</summary>
    public static NameLabelPrintMode NameLabelMode { get; private set; } = NameLabelPrintMode.Auto;

    /// <summary>TSPL TEXT 使用的印表機字型名稱。</summary>
    public static string NameLabelTsplFont { get; private set; } = "TSS24.BF2";

    /// <summary>姓名貼 BITMAP 渲染優先字型（Windows 字型家族名稱）。</summary>
    public static string NameLabelBitmapFontFamily { get; private set; } = "Microsoft JhengHei";

    /// <summary>姓名貼文字是否旋轉 180 度。</summary>
    public static bool NameLabelRotate180 { get; private set; } = false;

    /// <summary>TSPL 列印速度（ips）。0 表示不下指令，沿用機器預設。</summary>
    public static int TscSpeed { get; private set; } = 0;

    /// <summary>TSPL 濃度（1-15）。0 表示不下指令，沿用機器預設。</summary>
    public static int TscDensity { get; private set; } = 0;

    /// <summary>TSPL CODEPAGE（例如 UTF-8）。空字串表示不下指令。</summary>
    public static string TscCodePage { get; private set; } = string.Empty;

    /// <summary>TSPL CHARSET（例如 UTF-8）。空字串表示不下指令。</summary>
    public static string TscCharSet { get; private set; } = string.Empty;

    /// <summary>BITMAP 二值化門檻（0-255，越低越不容易把灰階印成黑點）。</summary>
    public static int BitmapThreshold { get; private set; } = 128;

    /// <summary>BITMAP 字形加粗像素半徑（0 表示不加粗）。</summary>
    public static int BitmapBoldPx { get; private set; } = 0;

    /// <summary>是否將送印前的 BITMAP 存檔（除錯用）。</summary>
    public static bool DebugSaveBitmap { get; private set; } = false;

    /// <summary>送印前 BITMAP 存檔路徑（除錯用）。</summary>
    public static string DebugBitmapPath { get; private set; } = @"C:\test_label.png";

    /// <summary>姓名貼模式：列印完成後是否送出裁刀指令。</summary>
    public static bool TscCutAfterPrint { get; private set; } = false;

    /// <summary>紙張感應模式：gap / blackmark / continuous。</summary>
    public static PaperSensorMode TscPaperSensorMode { get; private set; } = PaperSensorMode.Gap;

    /// <summary>黑標高度（毫米），僅 blackmark 模式使用。</summary>
    public static double TscBlackMarkMm { get; private set; } = 2;

    /// <summary>黑標感應後裁切前額外走紙步數（dots）。</summary>
    public static int TscBlackMarkPostFeedSteps { get; private set; } = 0;

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
            case "NAME_LABEL_COLUMNS":
            case "TSC_NAME_COLUMNS":
                if (int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var cols) && cols > 0)
                    NameLabelColumns = cols;
                break;
            case "NAME_LABEL_ROWS":
            case "TSC_NAME_ROWS":
                if (int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var rows) && rows > 0)
                    NameLabelRows = rows;
                break;
            case "NAME_LABEL_CELL_GAP_MM":
            case "TSC_NAME_GAP_MM":
                if (double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var cellGap) && cellGap >= 0)
                {
                    NameLabelCellGapMm = cellGap;
                    // 舊參數仍可一次同步控制欄距與列距。
                    NameLabelColumnGapMm = cellGap;
                    NameLabelRowGapMm = cellGap;
                }
                break;
            case "NAME_LABEL_COLUMN_GAP_MM":
            case "TSC_NAME_COLUMN_GAP_MM":
                if (double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var colGap) && colGap >= 0)
                    NameLabelColumnGapMm = colGap;
                break;
            case "NAME_LABEL_ROW_GAP_MM":
            case "TSC_NAME_ROW_GAP_MM":
                if (double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var rowGap) && rowGap >= 0)
                    NameLabelRowGapMm = rowGap;
                break;
            case "NAME_LABEL_GRID_HEIGHT_MM":
            case "TSC_NAME_GRID_HEIGHT_MM":
                if (double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var gridHeightMm) && gridHeightMm >= 0)
                    NameLabelGridHeightMm = gridHeightMm;
                break;
            case "NAME_LABEL_LAYOUT_SCALE":
            case "TSC_NAME_LAYOUT_SCALE":
                if (double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var layoutScale) && layoutScale > 0)
                    NameLabelLayoutScale = Math.Clamp(layoutScale, 0.001, 2.0);
                break;
            case "NAME_LABEL_CHAR_SPACING_PX":
            case "TSC_NAME_CHAR_SPACING_PX":
                if (int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var charSpacingPx))
                    NameLabelCharSpacingPx = Math.Clamp(charSpacingPx, -200, 200);
                break;
            case "NAME_LABEL_FIRST_COLUMN_OFFSET_X_PX":
            case "TSC_NAME_FIRST_COLUMN_OFFSET_X_PX":
                if (int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var firstColumnOffsetPx))
                    NameLabelFirstColumnOffsetXPx = Math.Clamp(firstColumnOffsetPx, -200, 200);
                break;
            case "NAME_LABEL_COLUMN_OFFSETS_X_PX":
            case "TSC_NAME_COLUMN_OFFSETS_X_PX":
                var parts = value.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
                var list = new List<int>(parts.Length);
                foreach (var part in parts)
                {
                    if (int.TryParse(part, NumberStyles.Integer, CultureInfo.InvariantCulture, out var offsetPx))
                        list.Add(Math.Clamp(offsetPx, -200, 200));
                }
                NameLabelColumnOffsetsXPx = list.ToArray();
                break;
            case "NAME_LABEL_PRINT_MODE":
            case "TSC_NAME_LABEL_PRINT_MODE":
                if (value.Equals("text", StringComparison.OrdinalIgnoreCase))
                    NameLabelMode = NameLabelPrintMode.Text;
                else if (value.Equals("bitmap", StringComparison.OrdinalIgnoreCase))
                    NameLabelMode = NameLabelPrintMode.Bitmap;
                else
                    NameLabelMode = NameLabelPrintMode.Auto;
                break;
            case "NAME_LABEL_TSPL_FONT":
            case "TSC_NAME_LABEL_TSPL_FONT":
                NameLabelTsplFont = value;
                break;
            case "NAME_LABEL_BITMAP_FONT_FAMILY":
            case "TSC_NAME_LABEL_BITMAP_FONT_FAMILY":
                NameLabelBitmapFontFamily = value;
                break;
            case "NAME_LABEL_ROTATE_180":
            case "TSC_NAME_LABEL_ROTATE_180":
                NameLabelRotate180 =
                    value.Equals("1", StringComparison.OrdinalIgnoreCase)
                    || value.Equals("true", StringComparison.OrdinalIgnoreCase)
                    || value.Equals("yes", StringComparison.OrdinalIgnoreCase)
                    || value.Equals("on", StringComparison.OrdinalIgnoreCase);
                break;
            case "TSC_SPEED":
                if (int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var speed) && speed >= 0)
                    TscSpeed = speed;
                break;
            case "TSC_DENSITY":
                if (int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var density) && density >= 0)
                    TscDensity = density;
                break;
            case "TSC_CODEPAGE":
                TscCodePage = value;
                break;
            case "TSC_CHARSET":
                TscCharSet = value;
                break;
            case "BITMAP_THRESHOLD":
            case "NAME_LABEL_BITMAP_THRESHOLD":
                if (int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var threshold))
                    BitmapThreshold = Math.Clamp(threshold, 0, 255);
                break;
            case "BITMAP_BOLD_PX":
            case "NAME_LABEL_BITMAP_BOLD_PX":
                if (int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var boldPx) && boldPx >= 0)
                    BitmapBoldPx = Math.Min(boldPx, 4);
                break;
            case "DEBUG_SAVE_BITMAP":
            case "NAME_LABEL_DEBUG_SAVE_BITMAP":
                DebugSaveBitmap =
                    value.Equals("1", StringComparison.OrdinalIgnoreCase)
                    || value.Equals("true", StringComparison.OrdinalIgnoreCase)
                    || value.Equals("yes", StringComparison.OrdinalIgnoreCase)
                    || value.Equals("on", StringComparison.OrdinalIgnoreCase);
                break;
            case "DEBUG_BITMAP_PATH":
            case "NAME_LABEL_DEBUG_BITMAP_PATH":
                DebugBitmapPath = value;
                break;
            case "TSC_CUT_AFTER_PRINT":
            case "LABEL_CUT_AFTER_PRINT":
            case "TSC_AUTO_CUT":
                TscCutAfterPrint =
                    value.Equals("1", StringComparison.OrdinalIgnoreCase)
                    || value.Equals("true", StringComparison.OrdinalIgnoreCase)
                    || value.Equals("yes", StringComparison.OrdinalIgnoreCase)
                    || value.Equals("on", StringComparison.OrdinalIgnoreCase);
                break;
            case "TSC_PAPER_SENSOR_MODE":
            case "TSC_MEDIA_TYPE":
            case "LABEL_SENSOR_MODE":
                if (value.Equals("blackmark", StringComparison.OrdinalIgnoreCase)
                    || value.Equals("black_mark", StringComparison.OrdinalIgnoreCase)
                    || value.Equals("black-mark", StringComparison.OrdinalIgnoreCase)
                    || value.Equals("mark", StringComparison.OrdinalIgnoreCase))
                {
                    TscPaperSensorMode = PaperSensorMode.BlackMark;
                }
                else if (value.Equals("continuous", StringComparison.OrdinalIgnoreCase)
                         || value.Equals("cont", StringComparison.OrdinalIgnoreCase))
                {
                    TscPaperSensorMode = PaperSensorMode.Continuous;
                }
                else
                {
                    TscPaperSensorMode = PaperSensorMode.Gap;
                }
                break;
            case "TSC_BLACK_MARK_MM":
            case "LABEL_BLACK_MARK_MM":
            case "TSC_BLINE_MM":
                if (double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var blineMm) && blineMm >= 0)
                    TscBlackMarkMm = blineMm;
                break;
            case "TSC_BLACK_MARK_POST_FEED_STEPS":
            case "LABEL_BLACK_MARK_POST_FEED_STEPS":
            case "TSC_BLACKMARK_POST_FEED_STEPS":
                if (int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var postFeedSteps))
                    TscBlackMarkPostFeedSteps = Math.Clamp(postFeedSteps, 0, 20000);
                break;
        }
    }
}
