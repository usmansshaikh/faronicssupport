# Enables Windows 11 feature update offer via Windows Update on Windows 10
# Run as Local System (Deep Freeze Cloud), silent/idempotent
# Logs to: %ProgramData%\DF-W11\w11-request.log

$ErrorActionPreference = 'Stop'
$logDir = Join-Path $env:ProgramData 'DF-W11'
$log = Join-Path $logDir 'w11-request.log'
New-Item -Path $logDir -ItemType Directory -Force | Out-Null

Function Write-Log {
    param([string]$msg)
    $stamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    "$stamp  $msg" | Tee-Object -FilePath $log -Append | Out-Null
}

# ---- Settings you can change ----
$TargetProduct = 'Windows 11'   # Must be exactly: Windows 11
$TargetRelease = '24H2'         # Example: 23H2 or 24H2
# ---------------------------------

Try {
    # Basic OS checks
    $os = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion'
    $prodName = $os.ProductName
    if ($prodName -match '^Windows 11') {
        Write-Log "Already on Windows 11 ($prodName). Nothing to do."
        exit 0
    }
    if ($prodName -match 'LTSC') {
        Write-Log "Detected LTSC ($prodName). Windows Update feature upgrades to Windows 11 are not offered on LTSC. Exiting."
        exit 0
    }
    if ($prodName -notmatch '^Windows 10') {
        Write-Log "Unsupported edition for this script: $prodName"
        exit 0
    }

    # Policy path
    $wuPolicyKey = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate'
    if (-not (Test-Path $wuPolicyKey)) {
        New-Item -Path $wuPolicyKey -Force | Out-Null
        Write-Log "Created policy key: $wuPolicyKey"
    }

    # Clear potential blocks / deferrals that prevent feature upgrades
    $deferralProps = @(
        'DeferFeatureUpdatesPeriodInDays',
        'PauseFeatureUpdatesStartTime',
        'PauseFeatureUpdatesEndTime',
        'PauseFeatureUpdates',
        'BranchReadinessLevel' # can lock channel and block upgrades
    )
    foreach ($p in $deferralProps) {
        if (Get-ItemProperty -Path $wuPolicyKey -Name $p -ErrorAction SilentlyContinue) {
            Remove-ItemProperty -Path $wuPolicyKey -Name $p -ErrorAction SilentlyContinue
            Write-Log "Removed blocking/deferral policy: $p"
        }
    }

    # Set Target Release (the key that tells WU which OS/version to offer)
    New-ItemProperty -Path $wuPolicyKey -Name 'TargetReleaseVersion' -Value 1 -PropertyType DWord -Force | Out-Null
    New-ItemProperty -Path $wuPolicyKey -Name 'ProductVersion' -Value $TargetProduct -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $wuPolicyKey -Name 'TargetReleaseVersionInfo' -Value $TargetRelease -PropertyType String -Force | Out-Null
    Write-Log "Configured TargetRelease: ProductVersion='$TargetProduct', TargetReleaseVersionInfo='$TargetRelease'."

    # Optional: ensure scanning is not blocked by a long-running pause for quality updates
    $wuUxKey = 'HKLM:\SOFTWARE\Microsoft\WindowsUpdate\UX\Settings'
    if (Test-Path $wuUxKey) {
        foreach ($name in 'PauseUpdatesExpiryTime','PauseUpdatesStartTime','PauseFeatureUpdatesStartTime','PauseFeatureUpdatesEndTime') {
            if (Get-ItemProperty -Path $wuUxKey -Name $name -ErrorAction SilentlyContinue) {
                Remove-ItemProperty -Path $wuUxKey -Name $name -ErrorAction SilentlyContinue
                Write-Log "Cleared UX pause value: $name"
            }
        }
    }

    # Restart Windows Update components (safe, idempotent)
    Write-Log "Restarting Windows Update services..."
    $services = 'wuauserv','bits','usosvc'
    foreach ($svc in $services) {
        Try { Stop-Service $svc -Force -ErrorAction SilentlyContinue } Catch {}
    }
    Start-Sleep -Seconds 2
    foreach ($svc in $services) {
        Try { Start-Service $svc -ErrorAction SilentlyContinue } Catch {}
    }

    # Trigger a detection scan so WU knows about the new target
    $uso = "$env:SystemRoot\System32\UsoClient.exe"
    if (Test-Path $uso) {
        & $uso StartScan        | Out-Null
        Start-Sleep -Seconds 2
        & $uso StartDownload    | Out-Null
        Write-Log "Triggered USO scan and download kickoff."
    } else {
        Write-Log "UsoClient.exe not found; skipping scan trigger."
    }

    Write-Log "Done. Windows Update should offer Windows 11 ($TargetRelease) at/after the next maintenance window if the device is eligible."
    exit 0
}
Catch {
    Write-Log "ERROR: $($_.Exception.Message)"
    exit 1
}
