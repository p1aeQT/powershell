# =====================================================
# P1AE Mod Analyzer - Remote Launcher
# Opens the beautiful index.html from this repo
# =====================================================

Clear-Host

$HtmlUrl = "https://raw.githubusercontent.com/p1aeQT/powershell/refs/heads/main/index.html"

Write-Host "Downloading the web interface..." -ForegroundColor Cyan

try {
    $HtmlContent = Invoke-RestMethod -Uri $HtmlUrl -UseBasicParsing -ErrorAction Stop

    # Save to a temporary HTML file
    $TempPath = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "P1AE-Mod-Analyzer.html")
    $HtmlContent | Out-File -FilePath $TempPath -Encoding UTF8 -Force

    Write-Host "✅ Download successful!" -ForegroundColor Green
    Write-Host "Opening analyzer in your default browser..." -ForegroundColor Cyan

    Start-Process $TempPath

    Write-Host ""
    Write-Host "You can scan .jar/.zip files or use the folder scanner." -ForegroundColor Gray

} catch {
    Write-Host "❌ Failed to download index.html" -ForegroundColor Red
    Write-Host "Please check that this file exists:" -ForegroundColor Yellow
    Write-Host "https://raw.githubusercontent.com/p1aeQT/powershell/refs/heads/main/index.html" -ForegroundColor White
    Write-Host ""
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Yellow
}

Write-Host ""
Write-Host "Press any key to exit..." -ForegroundColor DarkGray
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")