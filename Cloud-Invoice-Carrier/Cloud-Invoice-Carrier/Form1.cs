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
        public Form1()
        {
            InitializeComponent();
            InitWebViewAsync();
        }

        private async void InitWebViewAsync()
        {
            await webView21.EnsureCoreWebView2Async(null);

            // 使用輸出目錄路徑，發布時請將 載具生成器2.html 複製到程式同目錄（或設為 Content + 複製到輸出）
            var htmlPath = Path.Combine(Application.StartupPath, "載具生成器2.html");
            if (!File.Exists(htmlPath))
                htmlPath = @"C:\Users\user\Documents\GitHub\Cloud-Invoice-Carrier\Cloud-Invoice-Carrier\Cloud-Invoice-Carrier\載具生成器2.html"; // 開發時 fallback
            webView21.CoreWebView2.WebMessageReceived += WebView_WebMessageReceived;

            webView21.CoreWebView2.NavigationCompleted += (s, args) =>
            {
                if (!args.IsSuccess) return;
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
                }
                catch { }
            };

            webView21.CoreWebView2.Navigate(new Uri(htmlPath).AbsoluteUri);
        }

        // 用 nullable 避免 CS8618 警告
        private sealed class CarrierImageMessage
        {
            public string? type { get; set; }
            public string? fileName { get; set; }
            public string? dataUrl { get; set; }
        }

        private void WebView_WebMessageReceived(object? sender, CoreWebView2WebMessageReceivedEventArgs e)
        {
            try
            {
                var json = e.WebMessageAsJson;
                var msg = JsonSerializer.Deserialize<CarrierImageMessage>(json);
                if (msg == null || msg.type != "carrierImage" || string.IsNullOrEmpty(msg.dataUrl))
                    return;

                // dataUrl: "data:image/png;base64,AAAA..."
                var commaIndex = msg.dataUrl.IndexOf(',');
                var base64 = commaIndex >= 0 ? msg.dataUrl[(commaIndex + 1)..] : msg.dataUrl;
                var bytes = Convert.FromBase64String(base64);

                // 1. 存到 C:\test
                var folder = @"C:\test";
                Directory.CreateDirectory(folder);
                var fileName = string.IsNullOrWhiteSpace(msg.fileName)
                    ? "carrier-with-background.png"
                    : msg.fileName;
                var filePath = Path.Combine(folder, fileName);
                File.WriteAllBytes(filePath, bytes);

                // 2. 送給 HiTi 印表機
                PrintToHitiCs200e(bytes);
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

        private void PrintToHitiCs200e(byte[] imageBytes)
        {
            using var ms = new MemoryStream(imageBytes);
            using var img = Image.FromStream(ms);
            using var pd = new PrintDocument();

            // TODO: 改成「印表機與掃描器」中 HiTi CS-200e 的實際名稱
            pd.PrinterSettings.PrinterName = "HiTi CS-200e";

            if (!pd.PrinterSettings.IsValid)
            {
                MessageBox.Show("找不到印表機 HiTi CS-200e，請確認名稱是否正確。");
                return;
            }

            pd.DefaultPageSettings.Landscape = img.Width > img.Height;
            pd.DefaultPageSettings.Margins = new Margins(0, 0, 0, 0);

            pd.PrintPage += (s, e) =>
            {
                var bounds = e.PageBounds;
                e.Graphics?.DrawImage(img, bounds.X, bounds.Y, bounds.Width, bounds.Height);
            };

            pd.Print();
        }

        private void webView21_Click(object sender, EventArgs e)
        {

        }
    }
}
