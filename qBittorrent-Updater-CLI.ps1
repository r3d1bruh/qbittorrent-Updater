<#  Update-qBittorrent.ps1  — one file
    - Daily check at 03:30 + catch-up at startup
    - Robust SourceForge redirect handling
    - Silent install to custom directory
    - Logging + interactive pause when run manually
    - Version normalization (handles "v5.0.3")
#>

param(
  [string]$InstallDir = "E:\!Piracy\qBittorrent",   # <-- default path set for you
  [ValidateSet("standard","qt6_lt20")]
  [string]$Variant = "standard",
  [string]$TaskName = "qBittorrent Auto Update",
  [string]$At = "03:30",
  [switch]$NoSchedule
)

$ErrorActionPreference = "Stop"

# --- logging ---
$LogDir = "$env:ProgramData\qBittorrentUpdater"
$Log    = Join-Path $LogDir "run.log"
New-Item -ItemType Directory -Path $LogDir -Force | Out-Null
function Log($m){ "[$(Get-Date -Format s)] $m" | Out-File $Log -Append -Encoding utf8 }

function Normalize-Version([string]$ver) {
  if ($null -eq $ver -or $ver.Trim() -eq "") { return $null }
  ($ver -replace '^[vV]', '').Trim()
}

function Get-InstalledVersion {
  $exe = Join-Path $InstallDir "qbittorrent.exe"
  if (Test-Path $exe) { return (Get-Item $exe).VersionInfo.ProductVersion }
  foreach ($root in @(
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
  )) {
    $k = Get-ChildItem $root -ErrorAction SilentlyContinue |
      Where-Object { (Get-ItemProperty $_.PsPath -ErrorAction SilentlyContinue).DisplayName -like "qBittorrent*" } |
      Select-Object -First 1
    if ($k) {
      $v = (Get-ItemProperty $k.PsPath -ErrorAction SilentlyContinue).DisplayVersion
      if ($v) { return $v }
    }
  }
  return $null
}

function Get-LatestStableVersion {
  $html = Invoke-WebRequest -UseBasicParsing -Uri "https://www.qbittorrent.org/"
  $m = [regex]::Match($html.Content, "Latest:\s*v?(?<ver>\d+(?:\.\d+)+)")
  if ($m.Success) { return $m.Groups["ver"].Value }
  throw "Could not detect latest version from qbittorrent.org"
}

function Get-DownloadUrl([string]$ver,[string]$variant) {
  $file = if ($variant -eq "qt6_lt20") { "qbittorrent_${ver}_qt6_lt20_x64_setup.exe" }
          else { "qbittorrent_${ver}_x64_setup.exe" }
  "https://sourceforge.net/projects/qbittorrent/files/qbittorrent-win32/qbittorrent-$ver/$file/download"
}

function Resolve-Redirect([string]$url) {
  $final = $url
  for ($i=0; $i -lt 6; $i++) {
    $resp = Invoke-WebRequest -Uri $final -MaximumRedirection 0 -UseBasicParsing -ErrorAction SilentlyContinue
    if ($resp.StatusCode -ge 300 -and $resp.Headers.Location) {
      $final = $resp.Headers.Location
    } else { break }
  }
  $final
}

function Ensure-NotRunning {
  $p = Get-Process -Name "qbittorrent" -ErrorAction SilentlyContinue
  if ($p) { $p | Stop-Process -Force }
}

function Run-Installer([string]$url,[string]$targetDir) {
  $tmp = Join-Path $env:TEMP ("qbittorrent_" + [Guid]::NewGuid().ToString("N") + ".exe")
  try {
    $final = Resolve-Redirect $url
    Log "FinalURL=$final"
    Invoke-WebRequest -Uri $final -OutFile $tmp -UseBasicParsing -Headers @{ 'User-Agent' = 'Wget' }
    if ((Get-Item $tmp).Length -lt 5MB) { throw "Download too small; likely HTML not EXE." }

    $dirNoQuote = ($targetDir.TrimEnd('\'))
    $args = "/S /D=$dirNoQuote"  # NSIS: /S silent, /D must be last and unquoted

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $tmp
    $psi.Arguments = $args
    $psi.UseShellExecute = $true
    $psi.Verb = "runas"
    $p = [System.Diagnostics.Process]::Start($psi)
    $p.WaitForExit()
    $p.ExitCode
  } finally {
    Remove-Item $tmp -Force -ErrorAction SilentlyContinue
  }
}

function Ensure-ScheduledTask {
  param([string]$taskName,[string]$at)

  if ($NoSchedule) { return }
  $existing = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
  if ($existing) { Log "Task exists: $taskName"; return }

  $scriptPath = $MyInvocation.MyCommand.Path
  $argList = @(
    '-NoProfile',
    '-ExecutionPolicy', 'Bypass',
    '-File', ('"{0}"' -f $scriptPath),
    '-InstallDir', ('"{0}"' -f $InstallDir),
    '-Variant', $Variant
  )
  $exe = (Get-Command powershell.exe).Source
  $action  = New-ScheduledTaskAction -Execute $exe -Argument ($argList -join ' ')
  $time = [DateTime]::Parse($at)
  $triggerDaily   = New-ScheduledTaskTrigger -Daily -At $time
  $triggerStartup = New-ScheduledTaskTrigger -AtStartup
  $settings = New-ScheduledTaskSettingsSet `
               -StartWhenAvailable `
               -AllowStartIfOnBatteries `
               -DontStopIfGoingOnBatteries `
               -RunOnlyIfNetworkAvailable `
               -MultipleInstances IgnoreNew
  $principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
  Register-ScheduledTask -TaskName $taskName -Action $action -Trigger @($triggerDaily,$triggerStartup) `
    -Settings $settings -Principal $principal -Description "Checks qBittorrent version and auto-updates silently."
  Log "Task created: $taskName"
}

function Should-Pause {
  $isConsole = ($Host.Name -eq 'ConsoleHost') -or ($env:SESSIONNAME -like 'Console*')
  $isConsole -and ($env:USERNAME -ne 'SYSTEM')
}

# --- Main ---
try {
  Log "start"
  if (-not $NoSchedule) { Ensure-ScheduledTask -taskName $TaskName -at $At }

  $installedRaw = Get-InstalledVersion
  $latestRaw    = Get-LatestStableVersion
  $installedNorm = Normalize-Version $installedRaw
  $latestNorm    = Normalize-Version $latestRaw

  Write-Host "Installed: $installedRaw"
  Write-Host "Latest:    $latestNorm"
  Log "Installed=$installedRaw Latest=$latestNorm"

  if ($installedNorm -and ([version]$latestNorm -le [version]$installedNorm)) {
    Write-Host "Already up to date."
    Log "no-update"
    if (Should-Pause) { Write-Host "`nPress Enter to exit"; [void][Console]::ReadLine() }
    exit 0
  }

  Write-Host "Updating to $latestNorm ..."
  Log "update-start $latestNorm"
  Ensure-NotRunning
  $dl = Get-DownloadUrl -ver $latestNorm -variant $Variant
  Write-Host "Downloading: $dl"
  Log "DownloadURL=$dl"
  $code = Run-Installer -url $dl -targetDir $InstallDir
  Start-Sleep -Seconds 2

  $newRaw  = Get-InstalledVersion
  $newNorm = Normalize-Version $newRaw
  Log "ExitCode=$code NewVersion=$newNorm"

  if ($newNorm -and ([version]$newNorm -eq [version]$latestNorm)) {
    Write-Host "Update complete. Installed version: $newNorm"
    Log "success"
    if (Should-Pause) { Write-Host "`nPress Enter to exit"; [void][Console]::ReadLine() }
    exit 0
  } else {
    Write-Error "Update failed. ExitCode=$code InstalledAfter='$newRaw'"
    Log "fail"
    if (Should-Pause) { Write-Host "`nPress Enter to exit"; [void][Console]::ReadLine() }
    exit 1
  }
}
catch {
  Write-Error $_
  Log "error: $_"
  if (Should-Pause) { Write-Host "`nPress Enter to exit"; [void][Console]::ReadLine() }
  exit 1
}
