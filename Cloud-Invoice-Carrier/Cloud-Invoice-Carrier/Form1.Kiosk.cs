using System.Drawing;
using Microsoft.Web.WebView2.Core;

namespace Cloud_Invoice_Carrier;

public partial class Form1
{
    private const int WmSysCommand = 0x0112;
    private const int ScMinimize = 0xF020;
    private const int ScRestore = 0xF120;
    private const int WmContextMenu = 0x007B;

    private const string AntiPinchZoomDocumentScript = """
        (function () {
          if (window.__PINCH_ZOOM_GUARD_INSTALLED__) return;
          window.__PINCH_ZOOM_GUARD_INSTALLED__ = true;

          document.addEventListener('touchmove', function (e) {
            if (e.touches && e.touches.length >= 2) {
              e.preventDefault();
            }
          }, { passive: false });

          document.addEventListener('gesturestart', function (e) {
            e.preventDefault();
          }, { passive: false });

          document.addEventListener('gesturechange', function (e) {
            e.preventDefault();
          }, { passive: false });
        })();
        """;

    private const string KioskDocumentScript = """
        (function () {
          if (window.__KIOSK_GUARD_INSTALLED__) return;
          window.__KIOSK_GUARD_INSTALLED__ = true;

          var style = document.createElement('style');
          style.textContent = [
            'html, body { overflow: hidden; touch-action: manipulation; }',
            'html, body, body *:not(input):not(textarea) {',
            '  -webkit-user-select: none !important;',
            '  user-select: none !important;',
            '  -webkit-touch-callout: none !important;',
            '}',
            'input, textarea {',
            '  -webkit-user-select: text !important;',
            '  user-select: text !important;',
            '}',
            'img, video, svg { -webkit-user-drag: none; user-drag: none; }'
          ].join('\n');
          (document.head || document.documentElement).appendChild(style);

          document.addEventListener('contextmenu', function (e) {
            e.preventDefault();
            return false;
          }, true);

          document.addEventListener('dragstart', function (e) {
            e.preventDefault();
            return false;
          }, true);

          document.addEventListener('selectstart', function (e) {
            var tag = (e.target && e.target.tagName) ? e.target.tagName.toLowerCase() : '';
            if (tag !== 'input' && tag !== 'textarea') {
              e.preventDefault();
              return false;
            }
          }, true);

          document.addEventListener('wheel', function (e) {
            if (e.ctrlKey) {
              e.preventDefault();
              return false;
            }
          }, { passive: false });

          document.addEventListener('keydown', function (e) {
            if (e.ctrlKey && (e.key === '+' || e.key === '=' || e.key === '-')) {
              e.preventDefault();
            }
            if (e.ctrlKey && e.shiftKey && (e.key === 'I' || e.key === 'J')) {
              e.preventDefault();
            }
            if (e.ctrlKey && (e.key === 'U' || e.key === 'S')) {
              e.preventDefault();
            }
          }, true);
        })();
        """;

    private void ApplyKioskMode()
    {
        FormBorderStyle = FormBorderStyle.None;
        WindowState = FormWindowState.Maximized;
        TopMost = true;
        MaximizeBox = false;
        MinimizeBox = false;
        ControlBox = false;
        KeyPreview = true;

        webView21.AllowExternalDrop = false;

        Resize += (_, _) => EnsureKioskWindowState();
        Activated += (_, _) => EnsureKioskWindowState();
    }

    private void ApplyNormalWindowMode()
    {
        FormBorderStyle = FormBorderStyle.Sizable;
        TopMost = false;
        MaximizeBox = true;
        MinimizeBox = true;
        ControlBox = true;
        StartPosition = FormStartPosition.CenterScreen;
        WindowState = FormWindowState.Normal;

        var area = Screen.FromControl(this).WorkingArea;
        var width = Math.Min(1024, (int)(area.Width * 0.85));
        var height = Math.Min(900, (int)(area.Height * 0.85));
        ClientSize = new Size(width, height);
    }

    private void EnsureKioskWindowState()
    {
        if (WindowState == FormWindowState.Minimized)
            WindowState = FormWindowState.Maximized;
    }

    private static void ConfigureWebViewZoomLock(CoreWebView2 core)
    {
        core.Settings.IsPinchZoomEnabled = false;
        core.Settings.IsZoomControlEnabled = false;
    }

    private async Task ConfigureWebViewZoomLockAsync(CoreWebView2 core)
    {
        ConfigureWebViewZoomLock(core);
        await core.AddScriptToExecuteOnDocumentCreatedAsync(AntiPinchZoomDocumentScript);
    }

    private static void ConfigureWebViewKioskSettings(CoreWebView2 core)
    {
        core.Settings.AreDefaultContextMenusEnabled = false;
        core.Settings.AreBrowserAcceleratorKeysEnabled = false;
        core.Settings.AreDevToolsEnabled = false;
        core.Settings.IsStatusBarEnabled = false;
    }

    private async Task ConfigureWebViewKioskAsync(CoreWebView2 core)
    {
        ConfigureWebViewKioskSettings(core);
        await core.AddScriptToExecuteOnDocumentCreatedAsync(KioskDocumentScript);
    }

    protected override void WndProc(ref Message m)
    {
        if (!AppEnvConfig.KioskEnabled)
        {
            base.WndProc(ref m);
            return;
        }

        if (m.Msg == WmContextMenu)
            return;

        if (m.Msg == WmSysCommand)
        {
            var command = m.WParam.ToInt32() & 0xFFF0;
            if (command is ScMinimize or ScRestore)
                return;
        }

        base.WndProc(ref m);
    }
}
