#Requires -RunAsAdministrator
param([string]$User="Public")
$ErrorActionPreference="Stop"
$LoadedByScript=$false

try {
    $U=Get-LocalUser -Name $User -ErrorAction Stop
    $SID=$U.SID.Value
    $Profile=(Get-CimInstance Win32_UserProfile | Where-Object {$_.SID -eq $SID}).LocalPath
    if(-not $Profile){throw "Profile missing for '$User'."}
    $Root="Registry::HKEY_USERS\$SID"
    if(-not (Test-Path $Root)){
        & reg.exe load "HKU\$SID" "$Profile\NTUSER.DAT"|Out-Null
        if($LASTEXITCODE -ne 0){throw "Could not load user registry hive."}
        $LoadedByScript=$true
    }

    function Remove-Policy($Path,$Name) {
        Remove-ItemProperty $Path -Name $Name -ErrorAction SilentlyContinue
    }

    $System="$Root\Software\Microsoft\Windows\CurrentVersion\Policies\System"
    $Explorer="$Root\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
    $Network="$Root\Software\Policies\Microsoft\Windows\Network Connections"

    "DisableTaskMgr","DisableRegistryTools","DisableCMD" |
        ForEach-Object { Remove-Policy $System $_ }

    "NoControlPanel","HideSCANetwork","NoRun","NoWinKeys","SettingsPageVisibility" |
        ForEach-Object { Remove-Policy $Explorer $_ }

    "NC_LanProperties","NC_LanChangeProperties","NC_LanConnect" |
        ForEach-Object { Remove-Policy $Network $_ }

    "Library lockdown removed from '$User'. Sign out/restart for the change to fully apply."
}
finally {
    if($LoadedByScript){
        [gc]::Collect();[gc]::WaitForPendingFinalizers()
        & reg.exe unload "HKU\$SID"|Out-Null
    }
}
