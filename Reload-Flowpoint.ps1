# ============================
#  Flowpoint Local Sideload Helper
#  Works with Classic Outlook
# ============================

Write-Host "`n🚀 Reloading Flowpoint Add-in for Outlook..." -ForegroundColor Cyan

# ---- Configurable paths ----
$manifestPath = "C:\Users\dennis.hyre\dev\Flowpoint\manifest.xml"
$wefPath = "$env:LOCALAPPDATA\Microsoft\Office\16.0\Wef"
$outlookExe = "$env:ProgramFiles\Microsoft Office\root\Office16\OUTLOOK.EXE"

# ---- Close Outlook if running ----
Write-Host "Stopping Outlook if running..."
Get-Process outlook -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 2

# ---- Clear old sideload manifests ----
Write-Host "Clearing previous Flowpoint manifests from WEF..."
if (Test-Path $wefPath) {
    Get-ChildItem -Path $wefPath -Recurse -Include *Flowpoint*,*.xml | Remove-Item -Force -ErrorAction SilentlyContinue
}

# ---- Ensure WEF directory exists ----
Write-Host "Ensuring WEF folder exists..."
New-Item -ItemType Directory -Force -Path $wefPath | Out-Null

# ---- Copy new manifest ----
Write-Host "Copying latest manifest..."
Copy-Item $manifestPath -Destination "$wefPath\FlowpointManifest.xml" -Force

# ---- Launch Outlook ----
if (Test-Path $outlookExe) {
    Write-Host "Launching Outlook..." -ForegroundColor Green
    Start-Process $outlookExe
} else {
    Write-Host "⚠️ Could not find Outlook.exe. Please start Outlook manually." -ForegroundColor Yellow
}

Write-Host "`n✅ Flowpoint reloaded. Open an email and check the Home ribbon for 'Archive Emails'." -ForegroundColor Cyan
