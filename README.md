# NQA (Network Quality Accesser) (Windows)

## Install with npm

One-command install (global):

```powershell
npm install -g github:aGamingGod1234/net-quality-tray
```

Launch from anywhere:

```powershell
nqa-tray
```

CLI commands:

```powershell
nqa-tray start
nqa-tray install
nqa-tray status
nqa-tray open-folder
nqa-tray remove-startup
```

`NQA.exe` is a native Windows tray app that monitors real network quality:
- Download speed probe
- Upload speed probe
- Latency
- Jitter
- Packet loss
- Consistency over recent samples

## Files
- `NQA.exe` - native WinForms tray app
- `settings.json` - app configuration
- `assets\NetQualitySentinel.ico` - app/tray branding icon
- `NativeApp\*.cs` - native app source code
- `NetworkQualityTray.ps1` - legacy script (no longer required for normal use)

## Run
```powershell
Start-Process "C:\NetQualityTray\NQA.exe"
```

## Startup
Startup uses:
- Registry path: `HKCU\Software\Microsoft\Windows\CurrentVersion\Run`
- Value name: `NQA`
- Value: `"C:\NetQualityTray\NQA.exe"`

## Build Native App
Use .NET Framework `csc.exe`:
```powershell
& "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe" `
  /nologo /target:winexe `
  /out:C:\NetQualityTray\NQA.exe `
  /win32icon:C:\NetQualityTray\assets\NetQualitySentinel.ico `
  /reference:System.dll `
  /reference:System.Core.dll `
  /reference:System.Drawing.dll `
  /reference:System.Windows.Forms.dll `
  /reference:System.Net.Http.dll `
  /reference:System.Web.Extensions.dll `
  C:\NetQualityTray\NativeApp\*.cs
```

## Runtime Controls
- Left click or double click tray icon: opens the 60-second quality graph window.
- Right-click tray icon: opens a consistent tray context menu (graph, settings, probe now, pause/resume, exit).
- Settings include:
- quality score thresholds
- quality tier colors
- startup toggle
- pause/resume toggle
- probe now, status, and exit actions
