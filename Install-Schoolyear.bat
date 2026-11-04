<# : Batch launcher - runs the PowerShell below from this same file
@echo off
setlocal
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "$f='%~f0'; & ([scriptblock]::Create((Get-Content -LiteralPath $f -Raw)))"
exit /b %ERRORLEVEL%
#>

# ============================================================
# Install-Schoolyear.bat
# One run, any state: not installed, older, same or newer version.
#  - Already on the target version  -> nothing to do, exit 0
#  - Not installed                   -> install
#  - Other version(s) installed      -> stop app + service, uninstall, wait, install
# Waits for Windows Installer to be free and retries on 1618.
# Logs: C:\Windows\Temp\Schoolyear_Upgrade.log (+ msiexec verbose logs)
# ============================================================

$TargetVersion = '3.21.7'
$MsiName       = "schoolyear-exams-win-$TargetVersion.msi"
$MsiSources    = @(
    "\\172.19.130.99\systeem\$MsiName",
    "\\172.19.130.99\systeem\map\$MsiName"
)
$ForceReinstall = $false   # set to $true to reinstall even if already on the target version

$TempDir      = 'C:\Windows\Temp'
$LocalMsi     = Join-Path $TempDir $MsiName
$MainLog      = Join-Path $TempDir 'Schoolyear_Upgrade.log'
$UninstallLog = Join-Path $TempDir 'Schoolyear_Uninstall.log'
$InstallLog   = Join-Path $TempDir 'Schoolyear_Install.log'

$OkCodes = @(0, 1605, 1641, 3010)   # success, not installed, success+restart

function Log([string]$m) {
    $line = "[{0}] {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $m
    Write-Output $line
    Add-Content -Path $MainLog -Value $line -ErrorAction SilentlyContinue
}

function Get-SchoolyearInstalls {
    $paths = @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
    )
    Get-ItemProperty $paths -ErrorAction SilentlyContinue |
        Where-Object { $_.DisplayName -like '*Schoolyear*' -and $_.PSChildName -match '^\{[0-9A-Fa-f-]{36}\}$' } |
        Select-Object @{n='Code';e={$_.PSChildName}}, DisplayName, DisplayVersion
}

function Wait-MsiFree([int]$TimeoutSec = 600) {
    # Windows Installer holds Global\_MSIExecute while an install/uninstall is running
    $sw = [Diagnostics.Stopwatch]::StartNew()
    while ($sw.Elapsed.TotalSeconds -lt $TimeoutSec) {
        $m = $null
        try {
            if (-not [Threading.Mutex]::TryOpenExisting('Global\_MSIExecute', [ref]$m)) { return $true }
            $m.Dispose()
        } catch { }   # exists but not openable = treat as busy
        Start-Sleep -Seconds 5
    }
    return $false
}

function Invoke-Msi([string]$Arguments, [string]$What) {
    for ($try = 1; $try -le 5; $try++) {
        if (-not (Wait-MsiFree)) { Log "Windows Installer still busy after 10 minutes." }
        Log "$What (attempt $try): msiexec.exe $Arguments"
        $p = Start-Process -FilePath 'msiexec.exe' -ArgumentList $Arguments -Wait -PassThru -WindowStyle Hidden
        Log "$What exit code: $($p.ExitCode)"
        if ($p.ExitCode -ne 1618) { return $p.ExitCode }   # 1618 = another install in progress
        Start-Sleep -Seconds 30
    }
    return 1618
}

function Stop-Schoolyear {
    Get-Process -ErrorAction SilentlyContinue | Where-Object { ($_.Name -like '*schoolyear*') -or ($_.Name -eq 'updater' -and $_.Path -like '*Schoolyear*') } |
        ForEach-Object { Log "Stopping process $($_.Name) ($($_.Id))"; Stop-Process -Id $_.Id -Force -ErrorAction SilentlyContinue }
    Get-Service -ErrorAction SilentlyContinue | Where-Object { $_.Name -like '*schoolyear*' -or $_.DisplayName -like '*Schoolyear*' } |
        ForEach-Object {
            if ($_.Status -ne 'Stopped') {
                Log "Stopping service $($_.Name)"
                Stop-Service -Name $_.Name -Force -ErrorAction SilentlyContinue
                try { $_.WaitForStatus('Stopped', [TimeSpan]::FromSeconds(60)) } catch { Log "Service $($_.Name) did not stop within 60s" }
            }
        }
    Start-Sleep -Seconds 3
}

try {
    Log '============================================================'
    Log "Schoolyear install/upgrade to $TargetVersion on $env:COMPUTERNAME as $env:USERNAME"

    $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    if (-not $isAdmin) { Log 'ERROR: must run as Administrator or SYSTEM.'; exit 1 }

    # --- What is installed now? ---
    $installed = @(Get-SchoolyearInstalls)
    if ($installed.Count -eq 0) { Log 'Schoolyear is not installed.' }
    else { $installed | ForEach-Object { Log "Found: $($_.DisplayName) $($_.DisplayVersion) $($_.Code)" } }

    $onTarget = @($installed | Where-Object { $_.DisplayVersion -eq $TargetVersion })
    if (-not $ForceReinstall -and $installed.Count -gt 0 -and $onTarget.Count -eq $installed.Count) {
        Log "Already on $TargetVersion. Nothing to do."
        exit 0
    }

    # --- Get the MSI locally ---
    $src = $MsiSources | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
    if (-not $src) { Log "ERROR: MSI not found or share not reachable: $($MsiSources -join ' ; ')"; exit 2 }
    Log "Copying $src to $LocalMsi"
    Copy-Item -LiteralPath $src -Destination $LocalMsi -Force -ErrorAction Stop

    # --- Remove whatever is there ---
    if ($installed.Count -gt 0) {
        Stop-Schoolyear
        foreach ($i in $installed) {
            $code = Invoke-Msi "/x $($i.Code) /qn /norestart /L*v `"$UninstallLog`"" "Uninstall $($i.DisplayVersion)"
            if ($OkCodes -notcontains $code) { Log "ERROR: uninstall failed ($code). See $UninstallLog"; exit $code }
        }
        # Wait until the uninstall has really finished (registry entries gone, installer idle)
        $sw = [Diagnostics.Stopwatch]::StartNew()
        while (@(Get-SchoolyearInstalls).Count -gt 0 -and $sw.Elapsed.TotalSeconds -lt 120) { Start-Sleep -Seconds 5 }
        if (@(Get-SchoolyearInstalls).Count -gt 0) { Log 'WARNING: uninstall entries still present after 2 minutes; continuing.' }
        Wait-MsiFree | Out-Null
        Stop-Schoolyear   # in case anything restarted
    }

    # --- Install ---
    $code = Invoke-Msi "/i `"$LocalMsi`" /qn /norestart ALLUSERS=1 /L*v `"$InstallLog`"" "Install $TargetVersion"
    if ($OkCodes -notcontains $code -or $code -eq 1605) { Log "ERROR: install failed ($code). See $InstallLog"; exit $code }

    # --- Verify ---
    $after = @(Get-SchoolyearInstalls | Where-Object { $_.DisplayVersion -eq $TargetVersion })
    if ($after.Count -eq 0) { Log "ERROR: install reported $code but $TargetVersion is not registered."; exit 3 }
    Log "SUCCESS: Schoolyear $TargetVersion installed (exit $code)."
    if ($code -in 1641, 3010) { Log 'A restart is needed to finish.' }
    Remove-Item -LiteralPath $LocalMsi -Force -ErrorAction SilentlyContinue
    exit 0
}
catch {
    Log "ERROR: $($_.Exception.Message)"
    exit 1
}
