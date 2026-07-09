<#
  Ashaar Poetry - one-click installer for Word on Windows (desktop).

  Windows desktop Word loads sideloaded add-ins from a "trusted catalog"
  (a shared folder). This script downloads the manifest into a local folder,
  shares that folder, registers it as a trusted catalog, and clears the Office
  cache. It self-elevates because sharing a folder requires administrator rights.

  Run:  right-click this file > "Run with PowerShell"
  (or double-click Install-Ashaar.bat which launches it for you).
#>

$ErrorActionPreference = 'Stop'

$ManifestUrl = 'https://abdealikhurrum.github.io/ashaar.js-Office/manifest.prod.xml'
$FolderPath  = Join-Path $env:LOCALAPPDATA 'AshaarAddin'
$ShareName   = 'AshaarAddin'
$CatalogId   = 'a03d69e0-06d9-47c9-8995-06d3bdb5b198'  # fixed => re-running supersedes

# --- Self-elevate (folder sharing needs admin); keep the same user account ----
$isAdmin = ([Security.Principal.WindowsPrincipal] `
  [Security.Principal.WindowsIdentity]::GetCurrent()
  ).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
if (-not $isAdmin) {
  Write-Host 'Requesting administrator rights (needed to share the add-in folder)...'
  Start-Process powershell -Verb RunAs -ArgumentList `
    "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`""
  return
}

Write-Host '-- Ashaar Poetry - installing add-in for Word --' -ForegroundColor Cyan
Write-Host ''

# --- Close Word ---------------------------------------------------------------
Write-Host '-> Closing Word...'
Get-Process WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 2

# --- Download the manifest ----------------------------------------------------
Write-Host '-> Downloading manifest...'
New-Item -ItemType Directory -Force -Path $FolderPath | Out-Null
Invoke-WebRequest -Uri $ManifestUrl -OutFile (Join-Path $FolderPath 'manifest.prod.xml') -UseBasicParsing

# --- Share the folder (idempotent) --------------------------------------------
Write-Host '-> Sharing the add-in folder...'
if (-not (Get-SmbShare -Name $ShareName -ErrorAction SilentlyContinue)) {
  New-SmbShare -Name $ShareName -Path $FolderPath -FullAccess "$env:USERNAME" | Out-Null
}
$SharePath = "\\$env:COMPUTERNAME\$ShareName"

# --- Register the trusted catalog in the registry (HKCU) ----------------------
Write-Host '-> Registering the add-in catalog...'
$RegPath = "HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\$CatalogId"
New-Item -Path $RegPath -Force | Out-Null
Set-ItemProperty -Path $RegPath -Name 'Id'    -Value $CatalogId
Set-ItemProperty -Path $RegPath -Name 'Url'   -Value $SharePath
Set-ItemProperty -Path $RegPath -Name 'Flags' -Value 3 -Type DWord   # 3 = enabled + show in menu

# --- Clear the Office WebView cache -------------------------------------------
Write-Host '-> Clearing Word add-in cache...'
$Cache = Join-Path $env:LOCALAPPDATA 'Microsoft\Office\16.0\Wef'
if (Test-Path $Cache) { Remove-Item "$Cache\*" -Recurse -Force -ErrorAction SilentlyContinue }

Write-Host ''
Write-Host '[OK] Installed.' -ForegroundColor Green
Write-Host '   Open Word, then:  Insert > Add-ins > My Add-ins > SHARED FOLDER tab > Ashaar Poetry'
Write-Host '   (Restart Word first if it was open.)'
Write-Host ''
Read-Host 'Press Enter to close'
