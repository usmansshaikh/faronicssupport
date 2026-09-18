#Requires -RunAsAdministrator
param([string]$User="Public")
$ErrorActionPreference="Stop"

$LogDir = "$env:ProgramData\Faronics"
$Log = Join-Path $LogDir "LibraryLockdown.log"
New-Item $LogDir -ItemType Directory -Force | Out-Null

function Log([string]$Message) {
    "$(Get-Date -Format s) $Message" | Tee-Object -FilePath $Log -Append
}

$LoadedByScript = $false

try {
    Log "Starting library lockdown for '$User'."

    $U = Get-LocalUser -Name $User -ErrorAction Stop
    $SID = $U.SID.Value

    $IsAdmin = Get-LocalGroupMember -Group "Administrators" -ErrorAction SilentlyContinue |
        Where-Object { $_.SID -eq $U.SID }

    if ($IsAdmin) { throw "'$User' must not be an administrator." }

    $Profile = (Get-CimInstance Win32_UserProfile |
        Where-Object { $_.SID -eq $SID }).LocalPath

    if (-not $Profile) {
        throw "Windows profile missing. Sign into '$User' once before deployment."
    }

    $Root = "Registry::HKEY_USERS\$SID"

    if (-not (Test-Path $Root)) {
        & reg.exe load "HKU\$SID" "$Profile\NTUSER.DAT" | Out-Null
        if ($LASTEXITCODE -ne 0) { throw "Could not load the user registry hive." }
        $LoadedByScript = $true
    }

    function Set-Policy {
        param($Path,$Name,$Value,$Type="DWord")
        if (-not (Test-Path $Path)) { New-Item $Path -Force | Out-Null }
        New-ItemProperty $Path -Name $Name -Value $Value -PropertyType $Type -Force | Out-Null
    }

    $System   = "$Root\Software\Microsoft\Windows\CurrentVersion\Policies\System"
    $Explorer = "$Root\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
    $Network  = "$Root\Software\Policies\Microsoft\Windows\Network Connections"

    # Core public-user restrictions
    Set-Policy $System   "DisableTaskMgr"       1
    Set-Policy $System   "DisableRegistryTools" 1
    Set-Policy $System   "DisableCMD"           1
    Set-Policy $Explorer "NoControlPanel"       1
    Set-Policy $Explorer "HideSCANetwork"       1
    Set-Policy $Explorer "NoRun"                1
    Set-Policy $Explorer "NoWinKeys"            1

    # Network configuration restrictions
    Set-Policy $Network "NC_LanProperties"       0
    Set-Policy $Network "NC_LanChangeProperties" 0
    Set-Policy $Network "NC_LanConnect"          0

    # Hide sensitive Settings pages if Settings is made available later
    $Pages = "hide:network;network-status;network-wifi;network-ethernet;network-vpn;network-proxy;network-mobilehotspot;bluetooth;windowsupdate;delivery-optimization;yourinfo;emailandaccounts;signinoptions;otherusers"
    Set-Policy $Explorer "SettingsPageVisibility" $Pages "String"

    Log "SUCCESS: lockdown applied to '$User' ($SID)."
    exit 0
}
catch {
    Log "ERROR: $($_.Exception.Message)"
    exit 1
}
finally {
    if ($LoadedByScript) {
        [gc]::Collect()
        [gc]::WaitForPendingFinalizers()
        & reg.exe unload "HKU\$SID" | Out-Null
    }
}
