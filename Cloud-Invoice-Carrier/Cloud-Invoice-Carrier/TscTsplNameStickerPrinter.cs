using System.Runtime.InteropServices;
using System.Text;

namespace Cloud_Invoice_Carrier;

/// <summary>將中文文字轉成單色點陣並以 TSPL BITMAP 送印（適用多數 TSC TSPL 機種）。</summary>
internal static class TscTsplNameStickerPrinter
{
    public static void Print(string text, string windowsPrinterName, int dpi, double widthMm, double heightMm, double gapMm)
    {
        if (string.IsNullOrWhiteSpace(text))
            throw new ArgumentException("列印內容不可為空白。", nameof(text));

        var tspl = BuildTspl(text, dpi, widthMm, heightMm, gapMm);
        RawPrinterHelper.SendUtf8Text(windowsPrinterName, tspl, "NameSticker");
    }

    internal static string BuildTspl(string text, int dpi, double widthMm, double heightMm, double gapMm)
    {
        var wDots = Math.Max(8, (int)Math.Round(widthMm / 25.4 * dpi, MidpointRounding.AwayFromZero));
        var hDots = Math.Max(8, (int)Math.Round(heightMm / 25.4 * dpi, MidpointRounding.AwayFromZero));

        using var bmp = RenderLabelBitmap(text, wDots, hDots, dpi);
        var widthBytes = (bmp.Width + 7) / 8;
        var hex = ToTscBitmapHex(bmp);

        var sb = new StringBuilder(capacity: hex.Length + 256);
        sb.Append(FormattableString.Invariant($"SIZE {widthMm:0.###} mm,{heightMm:0.###} mm\r\n"));
        sb.Append(FormattableString.Invariant($"GAP {gapMm:0.###} mm,0 mm\r\n"));
        sb.Append("DIRECTION 1\r\n");
        sb.Append("REFERENCE 0,0\r\n");
        sb.Append("SET TEAR ON\r\n");
        sb.Append("CLS\r\n");
        sb.Append("BITMAP 0,0,");
        sb.Append(widthBytes);
        sb.Append(',');
        sb.Append(bmp.Height);
        sb.Append(",0,");
        sb.Append(hex);
        sb.Append("\r\nPRINT 1,1\r\n");
        return sb.ToString();
    }

    private static Bitmap RenderLabelBitmap(string text, int widthDots, int heightDots, int dpi)
    {
        var bmp = new Bitmap(widthDots, heightDots, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
        bmp.SetResolution(dpi, dpi);

        using var g = Graphics.FromImage(bmp);
        g.Clear(Color.White);

        var margin = Math.Max(2f, Math.Min(widthDots, heightDots) * 0.02f);
        var layout = new RectangleF(margin, margin, widthDots - margin * 2, heightDots - margin * 2);

        using var format = new StringFormat
        {
            Alignment = StringAlignment.Center,
            LineAlignment = StringAlignment.Center,
            Trimming = StringTrimming.EllipsisCharacter,
            FormatFlags = StringFormatFlags.LineLimit
        };

        var families = new[] { "Microsoft JhengHei", "微軟正黑體", "MingLiU", "新細明體", "PMingLiU", "細明體" };
        var maxStart = (float)Math.Min(layout.Width, layout.Height) * 0.6f;

        static Font? TryFont(string family, float sizePx)
        {
            try
            {
                return new Font(family, sizePx, FontStyle.Bold, GraphicsUnit.Pixel);
            }
            catch
            {
                return null;
            }
        }

        Font? best = null;
        var low = 8f;
        var high = Math.Max(low, maxStart);

        for (var iter = 0; iter < 36 && high - low > 0.35f; iter++)
        {
            var mid = (low + high) * 0.5f;
            Font? picked = null;
            foreach (var family in families)
            {
                picked = TryFont(family, mid);
                if (picked != null)
                    break;
            }

            if (picked == null)
            {
                high = mid - 0.5f;
                continue;
            }

            var measured = g.MeasureString(text, picked, layout.Size, format);
            var fits = measured.Width <= layout.Width + 0.5f && measured.Height <= layout.Height + 0.5f;
            if (fits)
            {
                best?.Dispose();
                best = picked;
                low = mid;
            }
            else
            {
                picked.Dispose();
                high = mid - 0.5f;
            }
        }

        if (best == null)
        {
            best = TryFont("Microsoft JhengHei", 12f)
                   ?? new Font(SystemFonts.DefaultFont.FontFamily, 12f, FontStyle.Bold, GraphicsUnit.Pixel);
        }

        try
        {
            g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAliasGridFit;
            using var brush = new SolidBrush(Color.Black);
            g.DrawString(text, best, brush, layout, format);
        }
        finally
        {
            best.Dispose();
        }

        return bmp;
    }

    private static string ToTscBitmapHex(Bitmap bmp)
    {
        var w = bmp.Width;
        var h = bmp.Height;
        var widthBytes = (w + 7) / 8;

        var rect = new Rectangle(0, 0, w, h);
        var data = bmp.LockBits(rect, System.Drawing.Imaging.ImageLockMode.ReadOnly, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
        try
        {
            var stride = data.Stride;
            var scan = data.Scan0;
            var outBuf = new byte[widthBytes * h];

            for (var y = 0; y < h; y++)
            {
                for (var xByte = 0; xByte < widthBytes; xByte++)
                {
                    byte b = 0;
                    for (var bit = 0; bit < 8; bit++)
                    {
                        var x = xByte * 8 + bit;
                        if (x >= w)
                            continue;

                        var offset = y * stride + x * 4;
                        var blue = Marshal.ReadByte(scan, offset);
                        var green = Marshal.ReadByte(scan, offset + 1);
                        var red = Marshal.ReadByte(scan, offset + 2);
                        var lum = (red * 299 + green * 587 + blue * 114) / 1000;
                        if (lum < 128)
                            b |= (byte)(0x80 >> bit);
                    }

                    outBuf[y * widthBytes + xByte] = b;
                }
            }

            return Convert.ToHexString(outBuf);
        }
        finally
        {
            bmp.UnlockBits(data);
        }
    }
}
