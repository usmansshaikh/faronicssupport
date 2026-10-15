# Admin check (required for remote execution via iex)
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Error "This script requires Administrator privileges."
    exit 1
}
<#
  Set-NeverSleep.ps1
  Sets Sleep, Hibernate and Hybrid Sleep to Never on all power plans (AC + DC),
  and turns hibernation off (removes hiberfil.sys, disables Fast Startup).
#>

$ErrorActionPreference = 'Stop'
$log = "$env:WINDIR\Temp\Set-NeverSleep.log"
function Write-Log($msg) { "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')  $msg" | Tee-Object -FilePath $log -Append }

Write-Log "=== Start on $env:COMPUTERNAME ==="

# Power setting GUIDs (SUB_SLEEP)
$SUB_SLEEP     = '238c9fa8-0aad-41ed-83f4-97be242c8f20'
$STANDBYIDLE   = '29f6c1db-86da-48c5-9fdb-f2b67b1f44da'   # Sleep after
$HIBERNATEIDLE = '9d7815a6-7ee4-497e-8888-515a05f02364'   # Hibernate after
$HYBRIDSLEEP   = '94ac6d29-73ce-41a6-809f-6363ba21b47e'   # Allow hybrid sleep

# Apply to every power scheme
$schemes = powercfg /list | Select-String -Pattern '([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})' |
           ForEach-Object { $_.Matches[0].Value }

foreach ($s in $schemes) {
    foreach ($setting in $STANDBYIDLE, $HIBERNATEIDLE, $HYBRIDSLEEP) {
        powercfg /setacvalueindex $s $SUB_SLEEP $setting 0 | Out-Null
        powercfg /setdcvalueindex $s $SUB_SLEEP $setting 0 | Out-Null
    }
    Write-Log "Scheme $s : Sleep=Never, Hibernate=Never, HybridSleep=Off"
}

# Re-apply the active scheme so the change takes effect immediately
$active = (powercfg /getactivescheme) -replace '.*:\s*([0-9a-f-]{36}).*', '$1'
powercfg /setactive $active
Write-Log "Active scheme re-applied: $active"

# Turn hibernation off completely
powercfg /hibernate off
Write-Log "Hibernation disabled"

# Verify
$q = powercfg /query $active $SUB_SLEEP
Write-Log ("Verify (active scheme):`n" + ($q -join "`n"))
Write-Log "=== Done ==="
