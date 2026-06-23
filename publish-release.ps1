# 打包成可在未安裝 .NET 的 Windows 電腦直接執行的單一 exe
# 輸出：publish\Cloud-Invoice-Carrier-win-x64\Cloud-Invoice-Carrier.exe
# 使用方式：複製 exe 到目標電腦即可執行（.env、carrier-layout.json 可選放在 exe 旁覆寫）

$ErrorActionPreference = "Stop"
$Project = Join-Path $PSScriptRoot "Cloud-Invoice-Carrier\Cloud-Invoice-Carrier\Cloud-Invoice-Carrier.csproj"
$Profile = "win-x64-selfcontained"

Write-Host "正在發佈（自包含 win-x64）..." -ForegroundColor Cyan
dotnet publish $Project -p:PublishProfile=$Profile

if ($LASTEXITCODE -ne 0) {
    Write-Host "發佈失敗。" -ForegroundColor Red
    exit $LASTEXITCODE
}

$OutDir = Join-Path $PSScriptRoot "publish\Cloud-Invoice-Carrier-win-x64"
Write-Host ""
Write-Host "完成！輸出目錄：" -ForegroundColor Green
Write-Host "  $OutDir"
Write-Host ""
Write-Host "請將 Cloud-Invoice-Carrier.exe 複製到目標電腦執行。" -ForegroundColor Yellow
Write-Host "可選：在 exe 旁放 .env、carrier-layout.json 覆寫預設設定。" -ForegroundColor Yellow
Write-Host "目標電腦需已安裝 Microsoft Edge WebView2 執行階段（Win11 通常已內建）。" -ForegroundColor Yellow
