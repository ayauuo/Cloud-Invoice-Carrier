using System;
using System.Drawing;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Collections.Generic;
using System.Text.Json;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;

namespace Cloud_Invoice_Carrier   // TODO: 這裡改成你專案的 namespace
{
    public partial class Form1 : Form
    {
        private static readonly JsonSerializerOptions WebMessageJsonOptions = new()
        {
            PropertyNameCaseInsensitive = true
        };

        public Form1()
        {
            InitializeComponent();
            InitWebViewAsync();
        }

        private async void InitWebViewAsync()
        {
            await webView21.EnsureCoreWebView2Async(null);

            AppEnvConfig.Load(Application.StartupPath);

            var htmlFile = AppEnvConfig.Mode == AppEnvConfig.AppMode.NameLabel
                ? "姓名貼鍵盤.html"
                : "載具生成器2.html";

            // 使用輸出目錄路徑；發布時請將 html 與 .env 放在 exe 同目錄
            var htmlPath = Path.Combine(Application.StartupPath, htmlFile);
            if (!File.Exists(htmlPath))
            {
                htmlPath = AppEnvConfig.Mode == AppEnvConfig.AppMode.NameLabel
                    ? @"C:\Users\user\Documents\GitHub\Cloud-Invoice-Carrier\Cloud-Invoice-Carrier\Cloud-Invoice-Carrier\姓名貼鍵盤.html"
                    : @"C:\Users\user\Documents\GitHub\Cloud-Invoice-Carrier\Cloud-Invoice-Carrier\Cloud-Invoice-Carrier\載具生成器2.html"; // 開發時 fallback
            }
            webView21.CoreWebView2.WebMessageReceived += WebView_WebMessageReceived;

            webView21.CoreWebView2.NavigationCompleted += (s, args) =>
            {
                if (!args.IsSuccess) return;
                if (AppEnvConfig.Mode == AppEnvConfig.AppMode.NameLabel)
                {
                    try
                    {
                        var dbPath = Path.Combine(Application.StartupPath, "jszhuyin", "database.data");
                        if (File.Exists(dbPath))
                        {
                            var base64 = Convert.ToBase64String(File.ReadAllBytes(dbPath));
                            webView21.CoreWebView2.PostWebMessageAsJson(
                                JsonSerializer.Serialize(new { type = "setJsZhuyinDatabase", base64 }));
                        }
                    }
                    catch
                    {
                        // 由前端顯示錯誤訊息，避免阻斷主流程
                    }
                    return;
                }

                var asm = Assembly.GetExecutingAssembly();
                // 載具背景（畫面預覽用）
                var previewNames = asm.GetManifestResourceNames()
                    .Where(n => n.Contains("picture") && n.Contains("載具背景"))
                    .OrderBy(n => n, StringComparer.Ordinal)
                    .ToList();
                var previewDataUrls = new List<string>();
                foreach (var name in previewNames)
                {
                    try
                    {
                        using var stream = asm.GetManifestResourceStream(name);
                        if (stream == null) continue;
                        using var ms = new MemoryStream();
                        stream.CopyTo(ms);
                        var base64 = Convert.ToBase64String(ms.ToArray());
                        previewDataUrls.Add("data:image/png;base64," + base64);
                    }
                    catch { }
                }

                // 實際背景（下載／列印用）
                var actualNames = asm.GetManifestResourceNames()
                    .Where(n => n.Contains("picture") && n.Contains("實際背景"))
                    .OrderBy(n => n, StringComparer.Ordinal)
                    .ToList();
                var actualDataUrls = new List<string>();
                foreach (var name in actualNames)
                {
                    try
                    {
                        using var stream = asm.GetManifestResourceStream(name);
                        if (stream == null) continue;
                        using var ms = new MemoryStream();
                        stream.CopyTo(ms);
                        var base64 = Convert.ToBase64String(ms.ToArray());
                        actualDataUrls.Add("data:image/png;base64," + base64);
                    }
                    catch { }
                }

                var titanDataUrls = new List<string>();
                var titanFront = BuildImageDataUrlFromRelativePath("picture/titan01.jpg");
                var titanBack = BuildImageDataUrlFromRelativePath("picture/titan02.jpg");
                if (!string.IsNullOrWhiteSpace(titanFront))
                    titanDataUrls.Add(titanFront);
                if (!string.IsNullOrWhiteSpace(titanBack))
                    titanDataUrls.Add(titanBack);
                try
                {
                    if (previewDataUrls.Count > 0)
                    {
                        webView21.CoreWebView2.PostWebMessageAsJson(JsonSerializer.Serialize(new { type = "setBackgroundImage", dataUrl = previewDataUrls[0] }));
                        if (previewDataUrls.Count > 1)
                            webView21.CoreWebView2.PostWebMessageAsJson(JsonSerializer.Serialize(new { type = "setBackgroundImages", dataUrls = previewDataUrls }));
                    }

                    if (actualDataUrls.Count > 0)
                    {
                        webView21.CoreWebView2.PostWebMessageAsJson(JsonSerializer.Serialize(new { type = "setActualBackgroundImages", dataUrls = actualDataUrls }));
                    }

                    if (titanDataUrls.Count > 0)
                    {
                        webView21.CoreWebView2.PostWebMessageAsJson(JsonSerializer.Serialize(new { type = "setTitanBackgroundImages", dataUrls = titanDataUrls }));
                    }
                }
                catch { }
            };

            webView21.CoreWebView2.Navigate(new Uri(htmlPath).AbsoluteUri);
        }

        // 用 nullable 避免 CS8618 警告
        private sealed class HostWebMessage
        {
            public string? type { get; set; }
            /// <summary>姓名貼模式：要列印的中文字（送 TSC TSPL）。</summary>
            public string? text { get; set; }
            public string? fileName { get; set; }
            public string? dataUrl { get; set; }
            public string? dataUrlBack { get; set; }
            /// <summary>例如 picture/001.jpg，相對於 Application.StartupPath（避免 WebView canvas 無法 toDataURL）</summary>
            public string? relativePath { get; set; }
            /// <summary>雙面列印時第 2 面圖檔（例如 picture/巨人載具02.jpg）。</summary>
            public string? relativePathBack { get; set; }
            public bool duplex { get; set; }
            /// <summary>標註橫版／直版（僅影響雙面第 2 頁是否轉 180°）；null 表示自動（高大於寬時轉 180°）。紙張一律直向。</summary>
            public bool? isLandscape { get; set; }
        }

        private static bool IsPathInsideDirectory(string directory, string candidatePath)
        {
            var root = Path.GetFullPath(directory.TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar);
            var full = Path.GetFullPath(candidatePath);
            return full.StartsWith(root, StringComparison.OrdinalIgnoreCase);
        }

        private static string ResolvePicturePath(string rel)
        {
            var normalized = rel.Replace('/', Path.DirectorySeparatorChar).TrimStart(Path.DirectorySeparatorChar);
            var nextToExe = Path.Combine(Application.StartupPath, normalized);
            if (File.Exists(nextToExe))
                return nextToExe;

            var fileName = Path.GetFileName(normalized);
            var devDir = @"C:\Users\user\Documents\GitHub\Cloud-Invoice-Carrier\Cloud-Invoice-Carrier\Cloud-Invoice-Carrier\picture";
            if (!string.IsNullOrWhiteSpace(fileName))
            {
                var devPath = Path.Combine(devDir, fileName);
                if (File.Exists(devPath))
                    return devPath;
            }

            return nextToExe;
        }

        private static string? BuildImageDataUrlFromRelativePath(string relativePath)
        {
            var path = ResolvePicturePath(relativePath);
            if (!File.Exists(path))
                return null;

            var ext = Path.GetExtension(path).ToLowerInvariant();
            var mime = ext switch
            {
                ".png" => "image/png",
                ".jpg" or ".jpeg" => "image/jpeg",
                ".webp" => "image/webp",
                ".gif" => "image/gif",
                _ => "application/octet-stream"
            };
            var base64 = Convert.ToBase64String(File.ReadAllBytes(path));
            return $"data:{mime};base64,{base64}";
        }

        private void WebView_WebMessageReceived(object? sender, CoreWebView2WebMessageReceivedEventArgs e)
        {
            try
            {
                var json = e.WebMessageAsJson;
                var msg = JsonSerializer.Deserialize<HostWebMessage>(json, WebMessageJsonOptions);
                if (msg == null || string.IsNullOrEmpty(msg.type))
                    return;

                if (msg.type.Equals("tscNameLabelPrint", StringComparison.OrdinalIgnoreCase))
                {
                    var t = (msg.text ?? string.Empty).Trim();
                    if (t.Length == 0)
                    {
                        MessageBox.Show("請先輸入要列印的中文字。", "姓名貼列印", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }

                    try
                    {
                        TscTsplNameStickerPrinter.Print(
                            t,
                            AppEnvConfig.TscWindowsPrinterName,
                            AppEnvConfig.TscDpi,
                            AppEnvConfig.LabelWidthMm,
                            AppEnvConfig.LabelHeightMm,
                            AppEnvConfig.LabelGapMm);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(
                            "送印至 TSC 時發生錯誤：\n" + ex.Message,
                            "姓名貼列印失敗",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error);
                    }

                    return;
                }

                // 本機圖檔直接由 C# 讀取（略過 HTML canvas，避免 file:// 造成 canvas 污染無法 toDataURL）
                if (msg.type.Equals("printLocalImageFile", StringComparison.OrdinalIgnoreCase))
                {
                    if (string.IsNullOrWhiteSpace(msg.relativePath))
                        return;

                    var relFront = msg.relativePath.Replace('/', Path.DirectorySeparatorChar).TrimStart(Path.DirectorySeparatorChar);
                    var frontPath = Path.GetFullPath(Path.Combine(Application.StartupPath, relFront));
                    if (!IsPathInsideDirectory(Application.StartupPath, frontPath))
                    {
                        MessageBox.Show("不允許的檔案路徑。", "列印");
                        return;
                    }

                    if (!File.Exists(frontPath))
                        frontPath = ResolvePicturePath(relFront);

                    if (!File.Exists(frontPath))
                    {
                        MessageBox.Show(
                            "找不到要列印的正面檔案：\n" + frontPath + "\n\n請將圖檔放在程式目錄下（與 exe 同層的相對路徑）。",
                            "列印圖檔",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Warning);
                        return;
                    }

                    byte[]? backBytes = null;
                    if (msg.duplex && !string.IsNullOrWhiteSpace(msg.relativePathBack))
                    {
                        var relBack = msg.relativePathBack.Replace('/', Path.DirectorySeparatorChar).TrimStart(Path.DirectorySeparatorChar);
                        var backPath = Path.GetFullPath(Path.Combine(Application.StartupPath, relBack));
                        if (!IsPathInsideDirectory(Application.StartupPath, backPath))
                        {
                            MessageBox.Show("不允許的反面檔案路徑。", "列印");
                            return;
                        }

                        if (!File.Exists(backPath))
                            backPath = ResolvePicturePath(relBack);

                        if (!File.Exists(backPath))
                        {
                            MessageBox.Show(
                                "找不到要列印的反面檔案：\n" + backPath + "\n\n請將圖檔放在程式目錄下（與 exe 同層的相對路徑）。",
                                "列印圖檔",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Warning);
                            return;
                        }

                        backBytes = File.ReadAllBytes(backPath);
                    }

                    var frontBytes = File.ReadAllBytes(frontPath);

                    var folder = @"C:\test";
                    Directory.CreateDirectory(folder);
                    var outName = string.IsNullOrWhiteSpace(msg.fileName) ? Path.GetFileName(frontPath) : msg.fileName;
                    File.WriteAllBytes(Path.Combine(folder, outName!), frontBytes);
                    if (backBytes != null)
                        File.WriteAllBytes(Path.Combine(folder, "back-" + outName), backBytes);

                    PrintToHitiCs200e(frontBytes, backBytes, msg.duplex, msg.isLandscape);
                    return;
                }

                if (msg.type != "carrierImage" || string.IsNullOrEmpty(msg.dataUrl))
                    return;

                // dataUrl: "data:image/png;base64,AAAA..."
                var commaIndex = msg.dataUrl.IndexOf(',');
                var base64 = commaIndex >= 0 ? msg.dataUrl[(commaIndex + 1)..] : msg.dataUrl;
                var bytesFromWeb = Convert.FromBase64String(base64);
                byte[]? bytesFromWebBack = null;
                if (msg.duplex && !string.IsNullOrWhiteSpace(msg.dataUrlBack))
                {
                    var backCommaIndex = msg.dataUrlBack.IndexOf(',');
                    var backBase64 = backCommaIndex >= 0 ? msg.dataUrlBack[(backCommaIndex + 1)..] : msg.dataUrlBack;
                    bytesFromWebBack = Convert.FromBase64String(backBase64);
                }

                // 1. 存到 C:\test
                var folder2 = @"C:\test";
                Directory.CreateDirectory(folder2);
                var fileName = string.IsNullOrWhiteSpace(msg.fileName)
                    ? "carrier-with-background.png"
                    : msg.fileName;
                var filePath = Path.Combine(folder2, fileName);
                File.WriteAllBytes(filePath, bytesFromWeb);
                if (bytesFromWebBack != null)
                    File.WriteAllBytes(Path.Combine(folder2, "back-" + fileName), bytesFromWebBack);

                // 2. 送 HiTi：若提供 dataUrlBack，雙面第 2 頁改用反面合成圖；否則維持兩面同圖。
                PrintToHitiCs200e(bytesFromWeb, bytesFromWebBack, msg.duplex, msg.isLandscape);
            }
            catch (Exception ex)
            {
                MessageBox.Show("處理載具圖片時發生錯誤：\n" + ex.Message);
            }
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            // 如果你現在是用 InitWebViewAsync() 在建構式呼叫，就可以先留空
            // 或者你想在載入時做事情，也可以寫在這裡
        }

        /// <summary>
        /// 在記憶體產生旋轉 180° 的點陣圖。HiTi 等驅動常忽略 PrintPage 的 Graphics 世界座標變換，改送已旋轉的像素較可靠。
        /// </summary>
        private static Bitmap CreateBitmapRotated180(Image source)
        {
            var bmp = new Bitmap(source.Width, source.Height);
            bmp.SetResolution(source.HorizontalResolution, source.VerticalResolution);
            using (var g = Graphics.FromImage(bmp))
            {
                g.TranslateTransform(source.Width / 2f, source.Height / 2f);
                g.RotateTransform(180);
                g.TranslateTransform(-source.Width / 2f, -source.Height / 2f);
                g.DrawImage(source, new Rectangle(0, 0, source.Width, source.Height));
            }

            return bmp;
        }

        /// <param name="isLandscape">
        /// 標註橫版（true）／直版（false），僅用於雙面時第 2 頁是否記憶體旋轉 180°；null 為自動（高大於寬時轉）。列印紙張固定直向，不設 Landscape。
        /// </param>
        /// <remarks>
        /// 當 backImageBytes 不為 null 且 duplex=true 時，第 2 頁使用 backImageBytes；
        /// 否則第 1、2 頁皆使用同一張圖（與既有行為相容）。
        /// </remarks>
        private void PrintToHitiCs200e(byte[] frontImageBytes, byte[]? backImageBytes, bool duplex, bool? isLandscape = null)
        {
            using var frontMs = new MemoryStream(frontImageBytes);
            using var frontImg = Image.FromStream(frontMs);
            Image? backImg = null;
            MemoryStream? backMs = null;
            if (duplex && backImageBytes != null)
            {
                backMs = new MemoryStream(backImageBytes);
                backImg = Image.FromStream(backMs);
            }

            using var pd = new PrintDocument();

            // TODO: 改成「印表機與掃描器」中 HiTi CS-200e 的實際名稱
            pd.PrinterSettings.PrinterName = "HiTi CS-200e";

            if (!pd.PrinterSettings.IsValid)
            {
                MessageBox.Show("找不到印表機 HiTi CS-200e，請確認名稱是否正確。");
                return;
            }

            // 不隨標註或圖素切換橫向紙張（僅改頁面方向、非旋轉圖素）；HiTi 一律直向頁面。
            pd.DefaultPageSettings.Landscape = false;
            pd.DefaultPageSettings.Margins = new Margins(0, 0, 0, 0);

            var duplexRotateBack180 = isLandscape switch
            {
                true => true,
                false => false,
                null => frontImg.Height > frontImg.Width
            };

            // 依使用者選擇設定雙面列印（HiTi 僅短邊翻面；符合 duplexRotateBack180 時第 2 頁送記憶體旋轉 180°）
            if (duplex)
            {
                if (pd.PrinterSettings.CanDuplex)
                    pd.PrinterSettings.Duplex = Duplex.Horizontal;
                else
                    pd.PrinterSettings.Duplex = Duplex.Simplex; // 若印表機不支援雙面，退回單面
            }
            else
            {
                pd.PrinterSettings.Duplex = Duplex.Simplex;
            }

            // 使用者選雙面時一律送兩頁（翻面模組／驅動常靠第二頁觸發）。
            var pageNumber = 1;
            var totalPages = duplex ? 2 : 1;

            pd.PrintPage += (s, e) =>
            {
                var bounds = e.PageBounds;
                var draw = (duplex && pageNumber == 2 && backImg != null) ? backImg : frontImg;

                var g = e.Graphics;
                if (g == null)
                {
                    e.HasMorePages = pageNumber < totalPages;
                    pageNumber++;
                    return;
                }

                var rotateBack180 = duplex && pageNumber == 2 && duplexRotateBack180;
                if (rotateBack180)
                {
                    using var rotated = CreateBitmapRotated180(draw);
                    g.DrawImage(rotated, bounds.X, bounds.Y, bounds.Width, bounds.Height);
                }
                else
                {
                    g.DrawImage(draw, bounds.X, bounds.Y, bounds.Width, bounds.Height);
                }

                e.HasMorePages = pageNumber < totalPages;
                pageNumber++;
            };

            try
            {
                pd.Print();
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "送印至 HiTi CS-200e 時發生錯誤：\n" + ex.Message,
                    "列印失敗",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            finally
            {
                backImg?.Dispose();
                backMs?.Dispose();
            }
        }

        /// <remarks>雙面載具時第 1、2 頁皆為同一張合成圖（實際背景＋條碼），與網頁輸出一致。</remarks>
        private void PrintToHitiCs200e(byte[] imageBytes, bool duplex, bool? isLandscape = null)
        {
            PrintToHitiCs200e(imageBytes, null, duplex, isLandscape);
        }

        private void webView21_Click(object sender, EventArgs e)
        {

        }
    }
}
