[CmdletBinding()]
param(
    [switch]$Restart
)

$ErrorActionPreference = "Stop"

$csc = "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
$outExe = "C:\NetQualityTray\NQA.exe"
$legacyExe = "C:\NetQualityTray\Net Quality Accesser.exe"
$icon = "C:\NetQualityTray\assets\NetQualitySentinel.ico"
$manifest = "C:\NetQualityTray\NativeApp\app.manifest"
$sources = "C:\NetQualityTray\NativeApp\*.cs"

if (-not (Test-Path $csc)) {
    throw "csc.exe not found at $csc"
}

$existing = Get-Process -Name NQA, NetQualitySentinel -ErrorAction SilentlyContinue
if ($existing) {
    $existing | Stop-Process -Force -ErrorAction SilentlyContinue
    Start-Sleep -Milliseconds 400
}

& $csc `
    /nologo `
    /target:winexe `
    /out:$outExe `
    /win32icon:$icon `
    /win32manifest:$manifest `
    /reference:System.dll `
    /reference:System.Core.dll `
    /reference:System.Drawing.dll `
    /reference:System.Windows.Forms.dll `
    /reference:System.Net.Http.dll `
    /reference:System.Web.Extensions.dll `
    $sources

if ($LASTEXITCODE -ne 0) {
    throw "Build failed."
}

Write-Host "Built: $outExe"

Copy-Item -Path $outExe -Destination $legacyExe -Force
Write-Host "Synced: $legacyExe"

if ($Restart) {
    Start-Process -FilePath $outExe
    Write-Host "Started: NQA.exe"
}
