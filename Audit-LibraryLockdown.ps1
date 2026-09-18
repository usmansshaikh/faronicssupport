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
    $System="$Root\Software\Microsoft\Windows\CurrentVersion\Policies\System"
    $Explorer="$Root\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
    $Checks=@(
        @($System,"DisableTaskMgr","Task Manager"),
        @($System,"DisableRegistryTools","Registry Editor"),
        @($System,"DisableCMD","Command Prompt"),
        @($Explorer,"NoControlPanel","Control Panel / Settings"),
        @($Explorer,"HideSCANetwork","Network tray controls"),
        @($Explorer,"NoRun","Run command"),
        @($Explorer,"NoWinKeys","Windows hotkeys")
    )
    foreach($C in $Checks){
        $V=(Get-ItemProperty $C[0] -Name $C[1] -ErrorAction SilentlyContinue).($C[1])
        "{0,-28} {1}" -f $C[2], $(if($V -eq 1){"LOCKED"}else{"OPEN"})
    }
}
finally {
    if($LoadedByScript){
        [gc]::Collect();[gc]::WaitForPendingFinalizers()
        & reg.exe unload "HKU\$SID"|Out-Null
    }
}
