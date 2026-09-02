using System;
using System.Drawing;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Collections.Generic;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;

namespace Cloud_Invoice_Carrier   // TODO: 這裡改成你專案的 namespace
{
    public partial class Form1 : Form
    {
        private const string AppVirtualHost = "app.local";
        private string _webContentRoot = AppPaths.ContentRoot;

        private static readonly JsonSerializerOptions WebMessageJsonOptions = new()
        {
            PropertyNameCaseInsensitive = true
        };

        private static readonly JsonSerializerOptions CarrierLayoutJsonOptions = new()
        {
            PropertyNameCaseInsensitive = true,
            ReadCommentHandling = JsonCommentHandling.Skip,
            AllowTrailingCommas = true
        };

        private sealed class CarrierLayoutConfigFile
        {
            [JsonPropertyName("version")]
            public int Version { get; set; }

            /// <summary>雙面翻面補償預設：auto / landscape / portrait。</summary>
            [JsonPropertyName("defaultDuplexFlipCompensation")]
            public string? DefaultDuplexFlipCompensation { get; set; }

            /// <summary>列印方式預設：simplex / duplex。</summary>
            [JsonPropertyName("defaultPrintDuplex")]
            public string? DefaultPrintDuplex { get; set; }

            [JsonPropertyName("templates")]
            public Dictionary<string, CarrierTemplateFile>? Templates { get; set; }
        }

        private sealed class CarrierTemplateFile
        {
            [JsonPropertyName("front")]
            public CarrierSideFile? Front { get; set; }

            [JsonPropertyName("back")]
            public CarrierSideFile? Back { get; set; }
        }

        private sealed class CarrierSideFile
        {
            [JsonPropertyName("previewBarcodeX")]
            public int? PreviewBarcodeX { get; set; }

            [JsonPropertyName("previewBarcodeY")]
            public int? PreviewBarcodeY { get; set; }

            [JsonPropertyName("printBarcodeX")]
            public int? PrintBarcodeX { get; set; }

            [JsonPropertyName("printBarcodeY")]
            public int? PrintBarcodeY { get; set; }
        }

        public Form1()
        {
            InitializeComponent();
            AppEnvConfig.Load(AppPaths.ExeDirectory);
            if (AppEnvConfig.KioskEnabled)
                ApplyKioskMode();
            else
                ApplyNormalWindowMode();
            InitWebViewAsync();
        }

        private async void InitWebViewAsync()
        {
            await webView21.EnsureCoreWebView2Async(null);
            await ConfigureWebViewZoomLockAsync(webView21.CoreWebView2);
            if (AppEnvConfig.KioskEnabled)
                await ConfigureWebViewKioskAsync(webView21.CoreWebView2);

            var layoutPath = AppPaths.FindFile("carrier-layout.json")
                ?? AppPaths.FindFile("carrier-layout.example.json");
            AppEnvConfig.ApplyCarrierLayoutDefaults(layoutPath);
            AppEnvConfig.Load(AppPaths.ExeDirectory);

            var htmlFile = AppEnvConfig.Mode == AppEnvConfig.AppMode.NameLabel
                ? "姓名貼鍵盤.html"
                : "載具生成器2.html";

            var contentRoot = AppPaths.ContentRoot;
            var htmlPath = AppPaths.FindFile(htmlFile) ?? Path.Combine(contentRoot, htmlFile);
            if (!File.Exists(htmlPath))
            {
                htmlPath = AppEnvConfig.Mode == AppEnvConfig.AppMode.NameLabel
                    ? @"C:\Users\user\Documents\GitHub\Cloud-Invoice-Carrier\Cloud-Invoice-Carrier\Cloud-Invoice-Carrier\姓名貼鍵盤.html"
                    : @"C:\Users\user\Documents\GitHub\Cloud-Invoice-Carrier\Cloud-Invoice-Carrier\Cloud-Invoice-Carrier\載具生成器2.html"; // 開發時 fallback
                contentRoot = Path.GetDirectoryName(htmlPath) ?? AppPaths.ContentRoot;
            }
            else
            {
                contentRoot = Path.GetDirectoryName(htmlPath) ?? AppPaths.ContentRoot;
            }

            _webContentRoot = contentRoot;

            webView21.CoreWebView2.SetVirtualHostNameToFolderMapping(
                AppVirtualHost,
                contentRoot,
                CoreWebView2HostResourceAccessKind.Allow);

            webView21.CoreWebView2.WebMessageReceived += WebView_WebMessageReceived;
            WireBillAcceptorForCarrierMode();
            EnsureBillAcceptorStartedAfterNavigation();

            webView21.CoreWebView2.NavigationCompleted += (_, args) =>
            {
                if (!args.IsSuccess) return;
                if (AppEnvConfig.Mode == AppEnvConfig.AppMode.NameLabel)
                {
                    _ = Task.Run(() =>
                    {
                        try
                        {
                            var dbPath = AppPaths.FindFile(Path.Combine("jszhuyin", "database.data"))
                                ?? AppPaths.CombineContent(Path.Combine("jszhuyin", "database.data"));
                            if (!File.Exists(dbPath))
                                return;

                            var base64 = Convert.ToBase64String(File.ReadAllBytes(dbPath));
                            BeginInvoke(() =>
                            {
                                if (webView21.CoreWebView2 == null) return;
                                try
                                {
                                    webView21.CoreWebView2.PostWebMessageAsJson(
                                        JsonSerializer.Serialize(new { type = "setJsZhuyinDatabase", base64 }));
                                }
                                catch { }
                            });
                        }
                        catch { }
                    });

                    try
                    {
                        var layout = new
                        {
                            type = "setNameLabelLayout",
                            dpi = AppEnvConfig.TscDpi,
                            widthMm = AppEnvConfig.LabelWidthMm,
                            heightMm = AppEnvConfig.LabelHeightMm,
                            previewWidthMm = AppEnvConfig.NameLabelPreviewWidthMm,
                            previewHeightMm = AppEnvConfig.NameLabelPreviewHeightMm,
                            columns = AppEnvConfig.NameLabelColumns,
                            rows = AppEnvConfig.NameLabelRows,
                            columnGapMm = AppEnvConfig.NameLabelColumnGapMm,
                            rowGapMm = AppEnvConfig.NameLabelRowGapMm,
                            gridHeightMm = AppEnvConfig.NameLabelGridHeightMm,
                            layoutScale = AppEnvConfig.NameLabelLayoutScale,
                            charSpacingPx = AppEnvConfig.NameLabelCharSpacingPx,
                            firstColumnOffsetXPx = AppEnvConfig.NameLabelFirstColumnOffsetXPx,
                            columnOffsetsXPx = AppEnvConfig.NameLabelColumnOffsetsXPx,
                            rotate180 = AppEnvConfig.NameLabelRotate180,
                            defaultFontFamily = AppEnvConfig.NameLabelBitmapFontFamily,
                            defaultText = AppEnvConfig.NameLabelDefaultText
                        };
                        webView21.CoreWebView2.PostWebMessageAsJson(JsonSerializer.Serialize(layout));
                    }
                    catch
                    {
                        // 預覽僅輔助用途，無設定不致命
                    }
                    return;
                }

                try
                {
                    PostCarrierBootstrapData(LoadCarrierBootstrapConfig());
                    EnsureBillAcceptorStartedAfterNavigation();
                    SyncBillAcceptorForIdlePage();
                }
                catch { }
            };

            var relativeHtml = Path.GetRelativePath(contentRoot, htmlPath).Replace('\\', '/');
            webView21.CoreWebView2.Navigate(BuildAppVirtualHostUri(relativeHtml));
        }

        private static string BuildAppVirtualHostUri(string relativePath)
        {
            var segments = relativePath.Split('/', StringSplitOptions.RemoveEmptyEntries);
            var encoded = string.Join("/", segments.Select(Uri.EscapeDataString));
            return $"https://{AppVirtualHost}/{encoded}";
        }

        private sealed class CarrierBootstrapConfig
        {
            public bool AllowBackTemplateSelection { get; init; }
            public bool PrintBarcodeOnBack { get; init; }
            public CarrierLayoutConfigFile? LayoutConfig { get; set; }
        }

        private CarrierBootstrapConfig LoadCarrierBootstrapConfig()
        {
            var config = new CarrierBootstrapConfig
            {
                AllowBackTemplateSelection = AppEnvConfig.CarrierAllowBackTemplateSelection,
                PrintBarcodeOnBack = AppEnvConfig.CarrierPrintBarcodeOnBack
            };

            try
            {
                var layoutPath = AppPaths.FindFile("carrier-layout.json")
                    ?? AppPaths.FindFile("carrier-layout.example.json");

                if (!string.IsNullOrWhiteSpace(layoutPath) && File.Exists(layoutPath))
                {
                    var json = File.ReadAllText(layoutPath);
                    config.LayoutConfig = JsonSerializer.Deserialize<CarrierLayoutConfigFile>(json, CarrierLayoutJsonOptions);
                }
            }
            catch { }

            return config;
        }

        private void PostCarrierBootstrapData(CarrierBootstrapConfig config)
        {
            if (config.LayoutConfig != null)
            {
                webView21.CoreWebView2.PostWebMessageAsJson(
                    JsonSerializer.Serialize(new { type = "setCarrierLayoutConfig", config = config.LayoutConfig }, WebMessageJsonOptions));
            }

            webView21.CoreWebView2.PostWebMessageAsJson(JsonSerializer.Serialize(new
            {
                type = "setCarrierTemplateOptions",
                allowBackTemplateSelection = config.AllowBackTemplateSelection,
                templateMode = AppEnvConfig.CarrierTemplateMode,
                printBarcodeOnBack = config.PrintBarcodeOnBack,
                printBarcodeEnabled = AppEnvConfig.CarrierPrintBarcodeEnabled,
                frontTemplateImages = CarrierTemplateImageResolver.ResolveHeadImages(_webContentRoot),
                backTemplateImages = CarrierTemplateImageResolver.ResolveBackImages(_webContentRoot)
            }));

            webView21.CoreWebView2.PostWebMessageAsJson(JsonSerializer.Serialize(new
            {
                type = "setPrintSaveOptions",
                savePrintImage = AppEnvConfig.SavePrintImage
            }));

            webView21.CoreWebView2.PostWebMessageAsJson(JsonSerializer.Serialize(new
            {
                type = "setBillAcceptorConfig",
                enabled = AppEnvConfig.BillAcceptorEnabled
            }));

            webView21.CoreWebView2.PostWebMessageAsJson(JsonSerializer.Serialize(new
            {
                type = "setCarrierUiOptions",
                showDuplexFlipCompensation = AppEnvConfig.CarrierShowDuplexFlipCompensation
            }));

            webView21.CoreWebView2.PostWebMessageAsJson(JsonSerializer.Serialize(new
            {
                type = "setKioskConfig",
                kioskEnabled = AppEnvConfig.KioskEnabled
            }));
        }

        // 用 nullable 避免 CS8618 警告
        private sealed class HostWebMessage
        {
            public string? type { get; set; }
            /// <summary>姓名貼模式：要列印的中文字（送 TSC TSPL）。</summary>
            public string? text { get; set; }
            /// <summary>姓名貼模式：要一並列印的載具條碼（CODE39）。</summary>
            public string? carrier { get; set; }
            /// <summary>姓名貼模式：本次 BITMAP 渲染要使用的 Windows 字型家族名稱。</summary>
            public string? fontFamily { get; set; }
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

        private static string ResolvePicturePath(string rel)
        {
            var normalized = rel.Replace('/', Path.DirectorySeparatorChar).TrimStart(Path.DirectorySeparatorChar);
            var resolved = AppPaths.FindFile(normalized);
            if (!string.IsNullOrWhiteSpace(resolved))
                return resolved;

            var fileName = Path.GetFileName(normalized);
            var devDir = @"C:\Users\user\Documents\GitHub\Cloud-Invoice-Carrier\Cloud-Invoice-Carrier\Cloud-Invoice-Carrier\picture";
            if (!string.IsNullOrWhiteSpace(fileName))
            {
                var devPath = Path.Combine(devDir, fileName);
                if (File.Exists(devPath))
                    return devPath;
            }

            return AppPaths.CombineContent(normalized);
        }

        private void WebView_WebMessageReceived(object? sender, CoreWebView2WebMessageReceivedEventArgs e)
        {
            try
            {
                // 優先用字串本體（HTML 端 postMessage(JSON.stringify(...))）
                string json;
                try
                {
                    json = e.TryGetWebMessageAsString();
                }
                catch
                {
                    json = e.WebMessageAsJson;
                }

                if (string.IsNullOrWhiteSpace(json))
                    json = e.WebMessageAsJson;

                if (TryHandleBillAcceptorWebMessage(json))
                    return;
                if (TryHandleHostRpc(json))
                    return;

                var msg = JsonSerializer.Deserialize<HostWebMessage>(json, WebMessageJsonOptions);
                if (msg == null || string.IsNullOrEmpty(msg.type))
                    return;

                if (msg.type.Equals("tscNameLabelPrint", StringComparison.OrdinalIgnoreCase))
                {
                    var t = (msg.text ?? string.Empty).Trim();
                    var carrier = (msg.carrier ?? string.Empty).Trim();
                    if (t.Length == 0 && carrier.Length == 0)
                    {
                        MessageBox.Show("請先輸入要列印的姓名或載具。", "姓名貼列印", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }

                    try
                    {
                        var fontFamily = string.IsNullOrWhiteSpace(msg.fontFamily)
                            ? AppEnvConfig.NameLabelBitmapFontFamily
                            : msg.fontFamily.Trim();
                        TscTsplNameStickerPrinter.Print(
                            t,
                            carrier,
                            AppEnvConfig.TscWindowsPrinterName,
                            AppEnvConfig.TscDpi,
                            AppEnvConfig.LabelWidthMm,
                            AppEnvConfig.LabelHeightMm,
                            AppEnvConfig.NameLabelPreviewWidthMm,
                            AppEnvConfig.NameLabelPreviewHeightMm,
                            AppEnvConfig.LabelGapMm,
                            AppEnvConfig.NameLabelColumns,
                            AppEnvConfig.NameLabelRows,
                            AppEnvConfig.NameLabelColumnGapMm,
                            AppEnvConfig.NameLabelRowGapMm,
                            AppEnvConfig.NameLabelGridHeightMm,
                            AppEnvConfig.NameLabelLayoutScale,
                            AppEnvConfig.NameLabelCharSpacingPx,
                            AppEnvConfig.NameLabelFirstColumnOffsetXPx,
                            AppEnvConfig.NameLabelColumnOffsetsXPx,
                            AppEnvConfig.NameLabelMode,
                            AppEnvConfig.NameLabelTsplFont,
                            fontFamily,
                            AppEnvConfig.NameLabelRotate180,
                            AppEnvConfig.TscSpeed,
                            AppEnvConfig.TscDensity,
                            AppEnvConfig.TscCodePage,
                            AppEnvConfig.TscCharSet,
                            AppEnvConfig.BitmapThreshold,
                            AppEnvConfig.BitmapBoldPx,
                            AppEnvConfig.DebugSaveBitmap,
                            AppEnvConfig.DebugBitmapPath,
                            AppEnvConfig.TscCutAfterPrint,
                            AppEnvConfig.TscPaperSensorMode,
                            AppEnvConfig.TscBlackMarkMm,
                            AppEnvConfig.TscBlackMarkPostFeedSteps);
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
                    var frontPath = Path.GetFullPath(ResolvePicturePath(relFront));
                    if (!AppPaths.IsUnderAppRoots(frontPath))
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
                        var backPath = Path.GetFullPath(ResolvePicturePath(relBack));
                        if (!AppPaths.IsUnderAppRoots(backPath))
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

                    SavePrintImageToDisk(
                        string.IsNullOrWhiteSpace(msg.fileName) ? Path.GetFileName(frontPath) : msg.fileName!,
                        frontBytes,
                        backBytes);

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

                var fileName = string.IsNullOrWhiteSpace(msg.fileName)
                    ? "carrier-with-background.png"
                    : msg.fileName;
                SavePrintImageToDisk(fileName, bytesFromWeb, bytesFromWebBack);

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
                AppendPrintDebug("Printer invalid: HiTi CS-200e");
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
                // 使用者明確要求雙面時，不允許退回單面避免誤印兩張單面。
                if (!pd.PrinterSettings.CanDuplex)
                {
                    AppendPrintDebug(
                        $"duplex requested=true, canDuplex=false, printer={pd.PrinterSettings.PrinterName}, appStart={AppPaths.ExeDirectory}");
                    MessageBox.Show(
                        "目前印表機或驅動未回報雙面能力，已取消列印。\n\n請確認：\n1) 使用正確的 HiTi CS-200e 印表機佇列\n2) 已安裝/啟用雙面模組與對應驅動\n3) Windows 印表機內容中的雙面選項可用",
                        "雙面列印不可用",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return;
                }

                pd.PrinterSettings.Duplex = Duplex.Horizontal;
            }
            else
            {
                pd.PrinterSettings.Duplex = Duplex.Simplex;
            }

            AppendPrintDebug(
                $"before print: duplexRequested={duplex}, canDuplex={pd.PrinterSettings.CanDuplex}, printer={pd.PrinterSettings.PrinterName}, printerDuplex={pd.PrinterSettings.Duplex}, defaultDuplex={pd.DefaultPageSettings.PrinterSettings.Duplex}, totalPages={(duplex ? 2 : 1)}");

            // 部分驅動會在逐頁送印前覆蓋頁面設定，這裡在每頁前再強制一次雙面/單面。
            pd.QueryPageSettings += (s, e) =>
            {
                e.PageSettings.PrinterSettings.Duplex = duplex ? Duplex.Horizontal : Duplex.Simplex;
                AppendPrintDebug(
                    $"query page settings: duplexRequested={duplex}, effectiveDuplex={e.PageSettings.PrinterSettings.Duplex}");
            };

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
                AppendPrintDebug("print completed without exception");
            }
            catch (Exception ex)
            {
                AppendPrintDebug("print exception: " + ex.Message);
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

        private static void SavePrintImageToDisk(string fileName, byte[] frontBytes, byte[]? backBytes = null)
        {
            if (!AppEnvConfig.SavePrintImage)
                return;

            try
            {
                var folder = AppEnvConfig.PrintSaveFolder;
                if (string.IsNullOrWhiteSpace(folder))
                    folder = @"C:\test";

                Directory.CreateDirectory(folder);
                File.WriteAllBytes(Path.Combine(folder, fileName), frontBytes);
                if (backBytes != null)
                    File.WriteAllBytes(Path.Combine(folder, "back-" + fileName), backBytes);
            }
            catch
            {
                // 存檔失敗不可中斷列印流程。
            }
        }

        private static void AppendPrintDebug(string message)
        {
            try
            {
                var folder = @"C:\test";
                Directory.CreateDirectory(folder);
                var line = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] {message}{Environment.NewLine}";
                File.AppendAllText(Path.Combine(folder, "print-debug.log"), line);
            }
            catch
            {
                // 除錯紀錄失敗不可中斷列印流程。
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
