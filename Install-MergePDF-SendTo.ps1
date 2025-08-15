<# Install-MergePDF-SendTo.ps1
   Richtet PDF-Merge via "Senden an" ein (qpdf-only).
   - Detection: [R]einstall / [U]ninstall / [E]xit
   - winget-Install von qpdf (wenn moeglich)
   - Legt C:\Tools\Merge-PDF.ps1 und C:\Tools\Merge-PDF.cmd an
   - Erstellt SendTo-Verknuepfung "PDFs zusammenfuehren"
   - Status im Terminal, Pause am Ende
#>

$ErrorActionPreference = 'Stop'

function W-Info($m){ Write-Host "[*] $m" -ForegroundColor Cyan }
function W-Ok($m){ Write-Host "[OK] $m" -ForegroundColor Green }
function W-Warn($m){ Write-Host "[!!] $m" -ForegroundColor Yellow }
function W-Err($m){ Write-Host "[ERR] $m" -ForegroundColor Red }

# Admincheck / Self-elevate
$principal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (-not $principal.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)) {
  W-Warn "Starte neu mit Administratorrechten..."
  Start-Process -FilePath "powershell.exe" -Verb RunAs -ArgumentList "-NoProfile","-ExecutionPolicy","Bypass","-File","`"$PSCommandPath`""
  return
}

# Pfade
$toolsDir = 'C:\Tools'
$ps1Path  = Join-Path $toolsDir 'Merge-PDF.ps1'
$cmdPath  = Join-Path $toolsDir 'Merge-PDF.cmd'
$sendTo   = [Environment]::GetFolderPath('SendTo')
$lnkPath  = Join-Path $sendTo 'PDFs zusammenfuehren.lnk'

# Inhalte: Wrapper + Merge-Script (robust, temp-output, exclude Merged_*.pdf)
$cmdContent = @"
@echo off
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "$ps1Path" %*
exit /b %ERRORLEVEL%
"@

$ps1Content = @'
# Merge-PDF.ps1 - qpdf-only, Clipboard, Auto-Detect, Logging (robust)
[CmdletBinding()]
param([Parameter(ValueFromRemainingArguments = $true)][string[]]$ArgsPaths)

$ErrorActionPreference = 'Stop'
$log = Join-Path $env:TEMP ("MergePDF_{0}.log" -f (Get-Date -Format yyyyMMdd_HHmmss))
Start-Transcript -Path $log -Append | Out-Null
Add-Type -AssemblyName System.Windows.Forms | Out-Null
$NL = [Environment]::NewLine

function Resolve-PdfList {
  param([string[]]$inputs)
  Write-Host ("Raw Args: " + ($inputs -join ' | '))
  $resolved = @()
  foreach ($p in $inputs) {
    if ([string]::IsNullOrWhiteSpace($p)) { continue }
    try { $rp = (Resolve-Path -LiteralPath $p).Path; if ($rp) { $resolved += $rp } } catch { }
  }
  Write-Host ("Resolved: " + ($resolved -join ' | '))

  function Is-Dir([string]$path) {
    try { if ([IO.Directory]::Exists($path)) { return $true }
          $it = Get-Item -LiteralPath $path -ErrorAction Stop; return [bool]$it.PSIsContainer } catch { return $false }
  }

  if ($resolved.Count -eq 1 -and (Is-Dir $resolved[0])) {
    $list = Get-ChildItem -LiteralPath $resolved[0] -Filter *.pdf -File |
            Where-Object { $_.BaseName -notlike 'Merged_*' } |
            Sort-Object Name | Select-Object -ExpandProperty FullName
    Write-Host ("Folder mode, PDFs: " + ($list -join ' | '))
    return $list
  }

  $files = $resolved | Where-Object { $_ -match '\.pdf$' -and (Test-Path -LiteralPath $_) }
  $list = $files | Sort-Object -Unique
  Write-Host ("File mode, PDFs: " + ($list -join ' | '))
  return $list
}

function Find-Qpdf {
  $candidates = @()
  $cmd = Get-Command qpdf.exe -ErrorAction SilentlyContinue
  if ($cmd) { $candidates += $cmd.Source }
  if ($PSScriptRoot) {
    $candidates += Join-Path $PSScriptRoot 'qpdf.exe'
    $candidates += Join-Path (Join-Path $PSScriptRoot 'bin') 'qpdf.exe'
  }
  $searchGlobs = @(
    "$env:ProgramFiles\qpdf*\bin\qpdf.exe",
    "$env:ProgramFiles(x86)\qpdf*\bin\qpdf.exe",
    "$env:ChocolateyInstall\lib\qpdf*\tools\qpdf.exe",
    "$env:LOCALAPPDATA\Programs\qpdf*\bin\qpdf.exe"
  ) | Where-Object { $_ }
  foreach ($g in $searchGlobs) {
    try { $hit = Get-ChildItem -Path $g -ErrorAction SilentlyContinue | Select-Object -First 1 -ExpandProperty FullName
          if ($hit) { $candidates += $hit } } catch { }
  }
  foreach ($hive in 'HKLM','HKCU') {
    $regPath = "$hive:SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\qpdf.exe"
    try { $ap = Get-ItemProperty -Path $regPath -ErrorAction SilentlyContinue
          if ($ap -and $ap.'(Default)') { $candidates += $ap.'(Default)' }
          elseif ($ap -and $ap.Path) { $candidates += $ap.Path } } catch { }
  }
  foreach ($c in ($candidates | Select-Object -Unique)) { if ($c -and (Test-Path -LiteralPath $c)) { return $c } }
  return $null
}

function Get-UniqueOutFile { param([string]$TargetDir,[string]$BaseName)
  $i=0; while ($true) { $n = ($i -eq 0) ? "$BaseName.pdf" : ("{0}_{1}.pdf" -f $BaseName,$i)
    $f = Join-Path $TargetDir $n; if (-not (Test-Path -LiteralPath $f)) { return $f }; $i++ } }

try {
  $files = Resolve-PdfList -inputs $ArgsPaths
  if (-not $files -or $files.Count -lt 2) { throw "Mindestens zwei PDF-Dateien noetig (oder ein Ordner mit >= 2 PDFs). Uebergeben: $($ArgsPaths -join ', ')" }

  $qpdf = Find-Qpdf
  if (-not $qpdf) {
    throw @"
qpdf wurde nicht gefunden.
Installationsoptionen:
  - winget install QPDF.QPDF
  - Portable qpdf.exe neben dieses Script legen (\bin oder direkt im Script-Ordner).
Log: $log
"@
  }

  $first = Get-Item -LiteralPath $files[0]
  $targetDir = $first.Directory.FullName
  $folder = Split-Path -Path $targetDir -Leaf
  $stamp = Get-Date -Format "yyyy-MM-dd_HHmm"
  $baseName = "Merged_{0}_{1}" -f $folder, $stamp
  $outFile = Get-UniqueOutFile -TargetDir $targetDir -BaseName $baseName

  Write-Host "qpdf: $qpdf"
  Write-Host ("Merging -> " + $outFile)

  # erst in TEMP schreiben, dann verschieben (vermeidet OneDrive/Locks)
  $tempOut = Join-Path $env:TEMP ("qpdf_merge_{0}.pdf" -f ([guid]::NewGuid()))
  $args = @('--empty','--pages') + $files + @('--', $tempOut)
  & $qpdf @args
  if ($LASTEXITCODE -ne 0) { throw "qpdf ExitCode: $LASTEXITCODE. Siehe Log: $log" }
  if (-not (Test-Path -LiteralPath $tempOut)) { throw "Zwischendatei nicht gefunden: $tempOut" }

  try { Move-Item -LiteralPath $tempOut -Destination $outFile -Force }
  catch { $outFile = Get-UniqueOutFile -TargetDir $targetDir -BaseName ($baseName + "_m"); Move-Item -LiteralPath $tempOut -Destination $outFile -Force }

  $cbOK = $false
  try { Set-Clipboard -Value $outFile -ErrorAction Stop; $cbOK=$true } catch {
    try { $clip = Get-Command clip.exe -ErrorAction SilentlyContinue; if ($clip){ $outFile | & $clip.Source; $cbOK=$true } } catch { } }
  $msg = "Fertig:" + $NL + $outFile + ($cbOK ? ($NL + "(Kopiert in die Zwischenablage)") : "")
  [System.Windows.Forms.MessageBox]::Show($msg,"PDF Merge",'OK','Information') | Out-Null
}
catch {
  [System.Windows.Forms.MessageBox]::Show("Fehler:" + $NL + $($_.Exception.Message) + $NL + $NL + "Log: $log","PDF Merge",'OK','Error') | Out-Null
}
finally { Stop-Transcript | Out-Null }
'@

# Detection
function Test-Installed {
  $files = (Test-Path $ps1Path) -and (Test-Path $cmdPath)
  $lnk   = Test-Path $lnkPath
  return ($files -and $lnk)
}

# qpdf via winget sicherstellen
function Ensure-Qpdf {
  param([string]$ToolsDir = 'C:\Tools')

  # qpdf schon vorhanden?
  $q = Get-Command qpdf.exe -ErrorAction SilentlyContinue
  if ($q) { W-Ok ("qpdf bereits verfuegbar: {0}" -f $q.Source); return }

  # winget robust finden (auch wenn PATH/Alias im Admin-Token fehlt)
  $wingetExe = $null
  $cands = @()
  $gc = Get-Command winget.exe -ErrorAction SilentlyContinue
  if ($gc) { $cands += $gc.Source }
  $cands += "$env:LOCALAPPDATA\Microsoft\WindowsApps\winget.exe"
  $cands += "$env:LOCALAPPDATA\Microsoft\Windows\WindowsApps\winget.exe"
  $cands += "$env:ProgramFiles\WindowsApps\Microsoft.DesktopAppInstaller*\winget.exe"  # evtl. nicht auflistbar

  foreach ($c in ($cands | Select-Object -Unique)) {
    try { if ($c -and (Test-Path -LiteralPath $c)) { $wingetExe = $c; break } } catch {}
  }

  if ($wingetExe) {
    W-Info "Installiere qpdf via winget..."
    $args = @('install','QPDF.QPDF',
              '--accept-package-agreements','--accept-source-agreements',
              '--silent','--disable-interactivity')
    $p = Start-Process -FilePath $wingetExe -ArgumentList $args -Wait -PassThru
    if ($p.ExitCode -eq 0) { W-Ok "qpdf installiert (winget)."; return }
    W-Warn ("winget ExitCode {0}. Wechsel auf portable Fallback." -f $p.ExitCode)
  } else {
    W-Warn "winget nicht gefunden (App Installer/Alias). Nutze portable Fallback."
  }

  # --- Portable Fallback von GitHub ---
  try {
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $api  = Invoke-WebRequest -UseBasicParsing -Uri 'https://api.github.com/repos/qpdf/qpdf/releases/latest'
    $rel  = $api.Content | ConvertFrom-Json
    $asset = $rel.assets | Where-Object {
      $_.name -match 'msvc64.*\.zip$' -or $_.name -match 'mingw64.*\.zip$'
    } | Select-Object -First 1
    if (-not $asset) { throw "Kein passendes qpdf Release-Asset gefunden." }

    $zip  = Join-Path $env:TEMP $asset.name
    W-Info ("Lade qpdf portable: {0}" -f $asset.name)
    Invoke-WebRequest -UseBasicParsing -Uri $asset.browser_download_url -OutFile $zip

    $dest = Join-Path $ToolsDir 'qpdf'
    if (Test-Path $dest) { Remove-Item $dest -Recurse -Force -ErrorAction SilentlyContinue }
    New-Item -ItemType Directory -Path $dest -Force | Out-Null
    Expand-Archive -LiteralPath $zip -DestinationPath $dest -Force

    $qpdfPath = Get-ChildItem -Path $dest -Recurse -Filter qpdf.exe -ErrorAction SilentlyContinue |
                Select-Object -First 1 -ExpandProperty FullName
    if (-not $qpdfPath) { throw "qpdf.exe nicht in ZIP gefunden." }

    W-Ok ("qpdf portable eingerichtet: {0}" -f $qpdfPath)
  }
  catch {
    W-Err ("Portable qpdf Setup fehlgeschlagen: {0}" -f $_.Exception.Message)
  }
}

# SendTo-Verknuepfung
function New-SendToShortcut {
  if (-not (Test-Path $sendTo)) { throw "SendTo-Verzeichnis nicht gefunden: $sendTo" }
  $shell = New-Object -ComObject WScript.Shell
  $sc = $shell.CreateShortcut($lnkPath)
  $sc.TargetPath = $cmdPath
  $sc.WorkingDirectory = $toolsDir
  $sc.IconLocation = 'imageres.dll,-5302'
  $sc.Description = 'Ausgewaehlte PDFs (oder Ordner) zu einer PDF zusammenfuehren'
  $sc.Arguments = ''
  $sc.Save()
}

function Remove-Installation {
  if (Test-Path $lnkPath) { Remove-Item $lnkPath -Force -ErrorAction SilentlyContinue }
  if (Test-Path $ps1Path) { Remove-Item $ps1Path -Force -ErrorAction SilentlyContinue }
  if (Test-Path $cmdPath) { Remove-Item $cmdPath -Force -ErrorAction SilentlyContinue }
}

# Interaktiver Flow
if (Test-Installed) {
  W-Info "Bereits installiert erkannt."
  Write-Host "Optionen: [R]einstall / [U]ninstall / [E]xit" -ForegroundColor Yellow
  switch ((Read-Host "Bitte Auswahl").ToUpperInvariant()) {
    'U' { W-Info "Deinstallation..."; Remove-Installation; W-Ok "Deinstallation abgeschlossen."; Read-Host "Weiter mit Enter..."; return }
    'R' { W-Info "Neuinstallation wird durchgefuehrt..." }
    default { W-Info "Abbruch."; Read-Host "Weiter mit Enter..."; return }
  }
}

# Installation
W-Info "Erzeuge $toolsDir..."
New-Item -Path $toolsDir -ItemType Directory -Force | Out-Null

W-Info "Schreibe $ps1Path..."
Set-Content -Path $ps1Path -Value $ps1Content -Encoding UTF8

W-Info "Schreibe $cmdPath..."
Set-Content -Path $cmdPath -Value $cmdContent -Encoding ASCII

W-Info "Erzeuge SendTo-Verknuepfung..."
New-SendToShortcut
W-Ok "SendTo: 'PDFs zusammenfuehren' eingerichtet."

Ensure-Qpdf

W-Ok "Installation abgeschlossen."
Write-Host "Benutzung: Dateien/Ordner markieren -> Rechtsklick -> Senden an -> 'PDFs zusammenfuehren'." -ForegroundColor Gray
Read-Host "Weiter mit Enter..."
