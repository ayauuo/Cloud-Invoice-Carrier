using System.Text;
using System.IO;

namespace Cloud_Invoice_Carrier;

/// <summary>姓名貼 TSPL 送印：英文優先 TEXT，中文自動回退 BITMAP。</summary>
internal static class TscTsplNameStickerPrinter
{
    static TscTsplNameStickerPrinter()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
    }

    public static void Print(
        string text,
        string windowsPrinterName,
        int dpi,
        double widthMm,
        double heightMm,
        double gapMm,
        int columns,
        int rows,
        double columnGapMm,
        double rowGapMm,
        double gridHeightMm,
        double layoutScale,
        int charSpacingPx,
        int firstColumnOffsetXPx,
        IReadOnlyList<int> columnOffsetsXPx,
        AppEnvConfig.NameLabelPrintMode printMode,
        string tsplFontName,
        string bitmapFontFamily,
        bool rotate180,
        int speed,
        int density,
        string codePage,
        string charSet,
        int bitmapThreshold,
        int bitmapBoldPx,
        bool debugSaveBitmap,
        string debugBitmapPath,
        bool cutAfterPrint,
        AppEnvConfig.PaperSensorMode paperSensorMode,
        double blackMarkMm,
        int blackMarkPostFeedSteps)
    {
        if (string.IsNullOrWhiteSpace(text))
            throw new ArgumentException("列印內容不可為空白。", nameof(text));

        var encoding = ResolveTsplEncoding(codePage, charSet);
        var payload = BuildTsplPayload(
            text, dpi, widthMm, heightMm, gapMm, columns, rows, columnGapMm, rowGapMm, gridHeightMm, layoutScale, charSpacingPx, firstColumnOffsetXPx, columnOffsetsXPx,
            printMode, tsplFontName, bitmapFontFamily, rotate180, speed, density, codePage, charSet, bitmapThreshold, bitmapBoldPx, debugSaveBitmap, debugBitmapPath, cutAfterPrint, paperSensorMode, blackMarkMm, blackMarkPostFeedSteps, encoding);
        RawPrinterHelper.SendBytes(windowsPrinterName, payload, "NameSticker");
    }

    private static Encoding ResolveTsplEncoding(string codePage, string charSet)
    {
        // 文字模式中文建議用 Big5（950）。未指定時維持 UTF-8。
        if (int.TryParse(codePage?.Trim(), out var cp) && cp > 0)
        {
            try { return Encoding.GetEncoding(cp); } catch { }
        }

        if (!string.IsNullOrWhiteSpace(codePage))
        {
            var cpName = codePage.Trim();
            if (cpName.Equals("BIG5", StringComparison.OrdinalIgnoreCase))
                return Encoding.GetEncoding(950);
            if (cpName.Equals("UTF-8", StringComparison.OrdinalIgnoreCase) || cpName.Equals("UTF8", StringComparison.OrdinalIgnoreCase))
                return Encoding.UTF8;
            try { return Encoding.GetEncoding(cpName); } catch { }
        }

        if (!string.IsNullOrWhiteSpace(charSet))
        {
            var cs = charSet.Trim();
            if (cs.Equals("BIG5", StringComparison.OrdinalIgnoreCase))
                return Encoding.GetEncoding(950);
            if (cs.Equals("UTF-8", StringComparison.OrdinalIgnoreCase) || cs.Equals("UTF8", StringComparison.OrdinalIgnoreCase))
                return Encoding.UTF8;
        }

        return Encoding.UTF8;
    }

    internal static byte[] BuildTsplPayload(
        string text,
        int dpi,
        double widthMm,
        double heightMm,
        double gapMm,
        int columns,
        int rows,
        double columnGapMm,
        double rowGapMm,
        double gridHeightMm,
        double layoutScale,
        int charSpacingPx,
        int firstColumnOffsetXPx,
        IReadOnlyList<int> columnOffsetsXPx,
        AppEnvConfig.NameLabelPrintMode printMode,
        string tsplFontName,
        string bitmapFontFamily,
        bool rotate180,
        int speed,
        int density,
        string codePage,
        string charSet,
        int bitmapThreshold,
        int bitmapBoldPx,
        bool debugSaveBitmap,
        string debugBitmapPath,
        bool cutAfterPrint,
        AppEnvConfig.PaperSensorMode paperSensorMode,
        double blackMarkMm,
        int blackMarkPostFeedSteps,
        Encoding encoding)
    {
        var wDots = Math.Max(32, (int)Math.Round(widthMm / 25.4 * dpi, MidpointRounding.AwayFromZero));
        var hDots = Math.Max(32, (int)Math.Round(heightMm / 25.4 * dpi, MidpointRounding.AwayFromZero));
        var cols = Math.Max(1, columns);
        var rowCount = Math.Max(1, rows);
        var scale = Math.Clamp(layoutScale, 0.001, 2.0);
        int GetColumnOffsetDots(int col)
        {
            if (columnOffsetsXPx != null && col >= 0 && col < columnOffsetsXPx.Count)
                return columnOffsetsXPx[col];
            if (col == 0)
                return firstColumnOffsetXPx;
            return 0;
        }
        var scaledColumnGapMm = columnGapMm * scale;
        var scaledRowGapMm = rowGapMm * scale;
        var colGapDots = Math.Max(0, (int)Math.Round(scaledColumnGapMm / 25.4 * dpi, MidpointRounding.AwayFromZero));
        var rowGapDots = Math.Max(0, (int)Math.Round(scaledRowGapMm / 25.4 * dpi, MidpointRounding.AwayFromZero));
        var marginDots = Math.Max(2, (int)Math.Round(Math.Min(wDots, hDots) * 0.02, MidpointRounding.AwayFromZero));
        var totalUsableH = Math.Max(16, hDots - marginDots * 2);
        var targetGridH = gridHeightMm > 0
            ? Math.Max(16, (int)Math.Round(gridHeightMm / 25.4 * dpi, MidpointRounding.AwayFromZero))
            : totalUsableH;
        var gridH = Math.Min(totalUsableH, targetGridH);
        // 版面靠下：保留上方空白，將下方空白壓到最小。
        var gridTop = marginDots + Math.Max(0, totalUsableH - gridH);
        var usableW = Math.Max(16, wDots - marginDots * 2 - colGapDots * (cols - 1));
        var usableH = Math.Max(16, gridH - rowGapDots * (rowCount - 1));
        var cellW = Math.Max(8, usableW / cols);
        var cellH = Math.Max(8, usableH / rowCount);
        var safeText = (text ?? string.Empty).Trim().Replace("\"", "'");
        var hasNonAscii = safeText.Any(ch => ch > 127);
        var fontName = string.IsNullOrWhiteSpace(tsplFontName) ? "TSS24.BF2" : tsplFontName.Trim().Replace("\"", "");
        var textRotation = rotate180 ? 180 : 0;

        var ms = new MemoryStream(capacity: 65536);
        void WriteCmd(string cmd)
        {
            var cmdBytes = encoding.GetBytes(cmd);
            ms.Write(cmdBytes, 0, cmdBytes.Length);
        }

        WriteCmd(FormattableString.Invariant($"SIZE {widthMm:0.###} mm,{heightMm:0.###} mm\r\n"));
        switch (paperSensorMode)
        {
            case AppEnvConfig.PaperSensorMode.BlackMark:
                WriteCmd(FormattableString.Invariant($"BLINE {Math.Max(0, blackMarkMm):0.###} mm,0 mm\r\n"));
                break;
            case AppEnvConfig.PaperSensorMode.Continuous:
                WriteCmd("GAP 0 mm,0 mm\r\n");
                break;
            default:
                WriteCmd(FormattableString.Invariant($"GAP {gapMm:0.###} mm,0 mm\r\n"));
                break;
        }
        if (speed > 0)
            WriteCmd(FormattableString.Invariant($"SPEED {speed}\r\n"));
        if (density > 0)
            WriteCmd(FormattableString.Invariant($"DENSITY {density}\r\n"));
        if (!string.IsNullOrWhiteSpace(codePage))
            WriteCmd($"CODEPAGE {codePage.Trim()}\r\n");
        if (!string.IsNullOrWhiteSpace(charSet))
            WriteCmd($"SET CHARSET {charSet.Trim()}\r\n");
        WriteCmd("DIRECTION 1\r\n");
        WriteCmd("REFERENCE 0,0\r\n");
        WriteCmd("SET TEAR ON\r\n");
        WriteCmd("CLS\r\n");

        var useTextMode = printMode switch
        {
            AppEnvConfig.NameLabelPrintMode.Text => true,
            AppEnvConfig.NameLabelPrintMode.Bitmap => false,
            _ => !hasNonAscii
        };

        if (useTextMode)
        {
            // 純英數符號：直接用 TEXT 命令，最清晰也最快。
            for (var row = 0; row < rowCount; row++)
            {
                for (var col = 0; col < cols; col++)
                {
                    var x = marginDots + col * (cellW + colGapDots) + 2 + GetColumnOffsetDots(col);
                    var y = gridTop + row * (cellH + rowGapDots) + 2;
                    WriteCmd($"TEXT {x},{y},\"{fontName}\",{textRotation},1,1,\"{safeText}\"\r\n");
                }
            }
        }
        else
        {
            // 中文或混合文字：改用 BITMAP，避免印表機端中文字型不支援造成空白。
            using var bmp = RenderLabelBitmap(safeText, wDots, hDots, dpi, cols, rowCount, colGapDots, rowGapDots, marginDots, gridTop, gridH, scale, charSpacingPx, bitmapFontFamily, firstColumnOffsetXPx, columnOffsetsXPx, rotate180);
            var widthBytes = bmp.Width / 8;
            if (debugSaveBitmap)
                SaveBitmapForDebug(bmp, debugBitmapPath);

            var bitmapData = ToTscBitmapData(bmp, bitmapThreshold, bitmapBoldPx);
            WriteCmd($"BITMAP 0,0,{widthBytes},{bmp.Height},0,");
            ms.Write(bitmapData, 0, bitmapData.Length);
            WriteCmd("\r\n");
        }

        WriteCmd("\r\nPRINT 1,1\r\n");
        if (paperSensorMode == AppEnvConfig.PaperSensorMode.BlackMark && blackMarkPostFeedSteps > 0)
            WriteCmd(FormattableString.Invariant($"FEED {blackMarkPostFeedSteps}\r\n"));
        if (cutAfterPrint)
            WriteCmd("CUT\r\n");
        return ms.ToArray();
    }

    private static Bitmap RenderLabelBitmap(
        string text,
        int widthDots,
        int heightDots,
        int dpi,
        int cols,
        int rowCount,
        int colGapDots,
        int rowGapDots,
        int marginDots,
        int gridTopDots,
        int gridHeightDots,
        double layoutScale,
        int charSpacingPx,
        string preferredFontFamily,
        int firstColumnOffsetXPx,
        IReadOnlyList<int> columnOffsetsXPx,
        bool rotate180)
    {
        // 先在真實寬度畫布排版，再貼到 32 對齊畫布，避免破壞版面比例。
        using var realBmp = new Bitmap(widthDots, heightDots, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
        realBmp.SetResolution(dpi, dpi);
        using (var g = Graphics.FromImage(realBmp))
        {
            g.Clear(Color.White);
            // 先以高品質灰階渲染字形，再交給後段 threshold 做 1-bit 二值化。
            g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
            g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
            g.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.HighQuality;
            g.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighQuality;
            g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAliasGridFit;

            int GetColumnOffsetPx(int col)
            {
                if (columnOffsetsXPx != null && col >= 0 && col < columnOffsetsXPx.Count)
                    return columnOffsetsXPx[col];
                if (col == 0)
                    return firstColumnOffsetXPx;
                return 0;
            }

            var availableWidth = Math.Max(8f, widthDots - marginDots * 2f - colGapDots * (cols - 1));
            var availableHeight = Math.Max(8f, gridHeightDots - rowGapDots * (rowCount - 1));
            var cellWidth = Math.Max(4f, availableWidth / cols);
            var cellHeight = Math.Max(4f, availableHeight / rowCount);
            var scale = (float)Math.Clamp(layoutScale, 0.001, 2.0);
            // 為字形預留內距，避免抗鋸齒後筆畫碰撞邊界被裁切。
            var padX = Math.Max(2f, cellWidth * 0.1f * scale);
            var padY = Math.Max(2f, cellHeight * 0.12f * scale);
            var innerWidth = Math.Max(2f, cellWidth - padX * 2f);
            var innerHeight = Math.Max(2f, cellHeight - padY * 2f);
            var textLines = (text ?? string.Empty)
                .Replace("\r\n", "\n")
                .Split('\n', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
            var useTwoLineLayout = textLines.Length >= 2;
            var lineGapPx = Math.Max(1f, 1f * scale); // 手動雙行行距，隨比例同步放大。
            var trackingPx = (float)charSpacingPx;
            var layoutFormat = new StringFormat
            {
                Alignment = StringAlignment.Center,
                LineAlignment = StringAlignment.Center,
                Trimming = StringTrimming.None,
                // 不裁切文字，避免因格高不足而整段看起來像空白。
                FormatFlags = StringFormatFlags.NoWrap | StringFormatFlags.NoClip
            };

            var preferred = (preferredFontFamily ?? string.Empty).Trim();
            var fallbackFamilies = new[] { "Microsoft JhengHei", "微軟正黑體", "MingLiU", "新細明體", "PMingLiU", "細明體" };
            var families = string.IsNullOrWhiteSpace(preferred)
                ? fallbackFamilies
                : (new[] { preferred }).Concat(fallbackFamilies).Distinct(StringComparer.OrdinalIgnoreCase).ToArray();

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
            // 保底字級，避免字體被縮到幾乎不可見（看起來像沒印）。
            var low = Math.Max(8f, 10f * scale);
            // 主要依寬度決定縮字；但加上高度上限保護，避免字完全被裁掉看不到。
            var widthDrivenHigh = innerWidth * (0.42f * scale);
            var heightSafetyHigh = innerHeight * 0.9f;
            var high = Math.Max(low, Math.Min(widthDrivenHigh, heightSafetyHigh));

            float MeasureTrackedWidth(string raw, Font font)
            {
                if (string.IsNullOrEmpty(raw))
                    return 0f;

                if (Math.Abs(trackingPx) < 0.01f)
                    return g.MeasureString(raw, font, new SizeF(innerWidth, innerHeight), layoutFormat).Width;

                var width = 0f;
                for (var i = 0; i < raw.Length; i++)
                {
                    var ch = raw[i].ToString();
                    var chSize = g.MeasureString(ch, font, new SizeF(4096, innerHeight), layoutFormat);
                    width += chSize.Width;
                    if (i < raw.Length - 1)
                        width += trackingPx;
                }
                return width;
            }

            void DrawTrackedString(string raw, Font font, Brush brush, RectangleF layout)
            {
                if (string.IsNullOrEmpty(raw))
                    return;

                if (Math.Abs(trackingPx) < 0.01f)
                {
                    g.DrawString(raw, font, brush, layout, layoutFormat);
                    return;
                }

                var textWidth = MeasureTrackedWidth(raw, font);
                var textHeight = font.GetHeight(g);
                var startX = layout.X + (layout.Width - textWidth) / 2f;
                var y = layout.Y + (layout.Height - textHeight) / 2f;

                for (var i = 0; i < raw.Length; i++)
                {
                    var ch = raw[i].ToString();
                    var chSize = g.MeasureString(ch, font, new SizeF(4096, layout.Height), layoutFormat);
                    var charRect = new RectangleF(startX, y, chSize.Width + 1f, textHeight + 1f);
                    g.DrawString(ch, font, brush, charRect, layoutFormat);
                    startX += chSize.Width + trackingPx;
                }
            }

            bool FitsInInnerBox(Font font)
            {
                if (!useTwoLineLayout)
                {
                    var measuredWidth = MeasureTrackedWidth(text ?? string.Empty, font);
                    return measuredWidth <= innerWidth * 0.95f;
                }

                var firstLine = textLines[0];
                var secondLine = textLines[1];
                var line1Width = MeasureTrackedWidth(firstLine, font);
                var line2Width = MeasureTrackedWidth(secondLine, font);
                var lineHeight = font.GetHeight(g);
                var totalHeight = lineHeight * 2f + lineGapPx;
                var maxLineWidth = Math.Max(line1Width, line2Width);
                return maxLineWidth <= innerWidth * 0.95f;
            }

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

                var fits = FitsInInnerBox(picked);
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

            // 若字串很長導致沒有任何「fit」結果，退到最小字級以維持單行顯示。
            best ??= TryFont("Microsoft JhengHei", low)
                     ?? new Font(SystemFonts.DefaultFont.FontFamily, low, FontStyle.Bold, GraphicsUnit.Pixel);
            try
            {
                using var brush = new SolidBrush(Color.Black);
                for (var row = 0; row < rowCount; row++)
                {
                    for (var col = 0; col < cols; col++)
                    {
                        var x = marginDots + col * (cellWidth + colGapDots) + GetColumnOffsetPx(col);
                        var y = gridTopDots + row * (cellHeight + rowGapDots);
                        var layout = new RectangleF(x + padX + 40f, y + padY - 75f, innerWidth, innerHeight);
                        if (!useTwoLineLayout)
                        {
                            DrawTrackedString(text ?? string.Empty, best, brush, layout);
                            continue;
                        }

                        var firstLine = textLines[0];
                        var secondLine = textLines[1];
                        var lineHeight = best.GetHeight(g);
                        var totalHeight = lineHeight * 2f + lineGapPx;
                        var startY = layout.Y + (layout.Height - totalHeight) / 2f;

                        var line1Rect = new RectangleF(layout.X, startY, layout.Width, lineHeight + 1f);
                        var line2Rect = new RectangleF(layout.X, startY + lineHeight + lineGapPx, layout.Width, lineHeight + 1f);
                        DrawTrackedString(firstLine, best, brush, line1Rect);
                        DrawTrackedString(secondLine, best, brush, line2Rect);
                    }
                }
            }
            finally
            {
                layoutFormat.Dispose();
                best.Dispose();
            }
        }

        // 部分機種 DMA 以 32-bit 對齊讀取，寬度補到 32 的倍數可避免行尾錯讀雜線。
        var alignedWidth = ((widthDots + 31) / 32) * 32;
        var alignedBmp = new Bitmap(alignedWidth, heightDots, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
        alignedBmp.SetResolution(dpi, dpi);
        using (var gAligned = Graphics.FromImage(alignedBmp))
        {
            gAligned.Clear(Color.White);
            gAligned.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.None;
            gAligned.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.NearestNeighbor;
            gAligned.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.Half;
            gAligned.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighSpeed;
            gAligned.DrawImageUnscaled(realBmp, 0, 0);
        }

        // 先排版完成再整張旋轉，能減少逐字旋轉帶來的鋸齒雜訊。
        if (rotate180)
            alignedBmp.RotateFlip(RotateFlipType.Rotate180FlipNone);

        return alignedBmp;
    }

    private static byte[] ToTscBitmapData(Bitmap bmp, int threshold, int boldPx)
    {
        var w = bmp.Width;
        var h = bmp.Height;
        var widthBytes = (w + 7) / 8;
        var outBuf = new byte[widthBytes * h];
        var th = Math.Clamp(threshold, 0, 255);
        var radius = Math.Max(0, boldPx);

        bool IsBlackAt(int xx, int yy)
        {
            if (xx < 0 || xx >= w || yy < 0 || yy >= h)
                return false;

            var pixel = bmp.GetPixel(xx, yy);
            var lum = (pixel.R * 299 + pixel.G * 587 + pixel.B * 114) / 1000;
            return lum < th;
        }

        for (var y = 0; y < h; y++)
        {
            for (var x = 0; x < w; x++)
            {
                var isBlack = IsBlackAt(x, y);
                if (!isBlack && radius > 0)
                {
                    for (var dy = -radius; dy <= radius && !isBlack; dy++)
                    {
                        for (var dx = -radius; dx <= radius; dx++)
                        {
                            if (dx == 0 && dy == 0)
                                continue;
                            if (IsBlackAt(x + dx, y + dy))
                            {
                                isBlack = true;
                                break;
                            }
                        }
                    }
                }

                if (!isBlack)
                    continue;

                var byteIndex = y * widthBytes + (x / 8);
                var bitIndex = 7 - (x % 8); // TSPL: MSB first
                outBuf[byteIndex] |= (byte)(1 << bitIndex);
            }
        }

        // 某些機種的 BITMAP 位元極性相反：0=黑、1=白，需整體反相。
        for (var i = 0; i < outBuf.Length; i++)
            outBuf[i] = (byte)~outBuf[i];

        return outBuf;
    }

    private static void SaveBitmapForDebug(Bitmap bmp, string savePath)
    {
        try
        {
            var path = string.IsNullOrWhiteSpace(savePath) ? @"C:\test_label.png" : savePath.Trim();
            var dir = Path.GetDirectoryName(path);
            if (!string.IsNullOrWhiteSpace(dir))
                Directory.CreateDirectory(dir);
            bmp.Save(path, System.Drawing.Imaging.ImageFormat.Png);
        }
        catch
        {
            // 除錯存檔不可阻斷主流程。
        }
    }
}
