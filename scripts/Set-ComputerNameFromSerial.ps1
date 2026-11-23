<#
  Post-imaging rename and iCAM client configuration
  Run as a Deep Freeze Cloud post-imaging task (runs as SYSTEM).
  1. Reads the BIOS serial number.
  2. Renames the PC from the table below.
  3. On library customer PCs, points the iCAM Cafe Client and Print Client
     at the site's admin PC (same site code, ending -201).
  4. Restarts so the new name takes effect.
  Log: C:\Windows\Temp\PostImage-Rename.log
  To add or change devices, edit the table below (Role: CLIENT, ADMIN, PUBLIC or SKIP).
#>

$LogFile = 'C:\Windows\Temp\PostImage-Rename.log'
function Write-Log($Msg) {
    $line = "{0}  {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Msg
    Add-Content -Path $LogFile -Value $line
    Write-Output $line
}

$DeviceTable = @'
Serial,Name,Role,Location
GM1B9L5X,LBE-WV-L279-201,ADMIN,ENFIELD TOWN - GROUND FLOOR
GM1B9L60,LBE-WV-L279-001,CLIENT,ENFIELD TOWN - GROUND FLOOR
GM1B9L61,LBE-WV-L279-002,CLIENT,ENFIELD TOWN - GROUND FLOOR
GM1B9L6F,LBE-WV-L279-003,CLIENT,ENFIELD TOWN - GROUND FLOOR
GM1B9L6H,LBE-WV-L279-004,CLIENT,ENFIELD TOWN - GROUND FLOOR
GM1B9L6V,LBE-WV-L279-005,CLIENT,ENFIELD TOWN - GROUND FLOOR
GM1B9L70,LBE-WV-L279-006,CLIENT,ENFIELD TOWN - GROUND FLOOR
GM1B9L71,LBE-WV-L279-007,CLIENT,ENFIELD TOWN - GROUND FLOOR
GM1B9L77,LBE-WV-L279-008,CLIENT,ENFIELD TOWN - GROUND FLOOR
GM1B9L7G,LBE-WV-L279-009,CLIENT,ENFIELD TOWN - GROUND FLOOR
GM1B9LA5,LBE-DT-000261,PUBLIC,Enfield Town - Ground floor Community Hub
GM1B9LA6,LBE-DT-000247,PUBLIC,Enfield Town - Ground floor Community Hub
GM1B9L7S,LBE-WV-L280-201,ADMIN,ENFIELD TOWN FIRST FLOOR
GM1B9L8B,LBE-WV-L280-001,CLIENT,ENFIELD TOWN FIRST FLOOR
GM1B9L8V,LBE-WV-L280-002,CLIENT,ENFIELD TOWN FIRST FLOOR
GM1B9L97,LBE-WV-L280-003,CLIENT,ENFIELD TOWN FIRST FLOOR
GM1B9L9D,LBE-WV-L280-004,CLIENT,ENFIELD TOWN FIRST FLOOR
GM1B9L9E,LBE-WV-L280-005,CLIENT,ENFIELD TOWN FIRST FLOOR
GM1B9L9M,LBE-WV-L280-006,CLIENT,ENFIELD TOWN FIRST FLOOR
GM1B9L9S,LBE-WV-L280-007,CLIENT,ENFIELD TOWN FIRST FLOOR
GM1B9L9Z,LBE-WV-L280-008,CLIENT,ENFIELD TOWN FIRST FLOOR
GM1B9LA3,LBE-WV-L280-009,CLIENT,ENFIELD TOWN FIRST FLOOR
GM1B9LA4,LBE-WV-L280-010,CLIENT,ENFIELD TOWN FIRST FLOOR
GM1B9L7F,LBE-WV-L280-011,CLIENT,ENFIELD TOWN FIRST FLOOR
GM1B9L9K,LBE-WV-L278-201,ADMIN,Oakwood
GM1B9L63,LBE-WV-L278-001,CLIENT,Oakwood
GM1B9L5R,LBE-WV-L278-002,CLIENT,Oakwood
GM1B9L5S,LBE-WV-L278-003,CLIENT,Oakwood
GM1B9L5W,LBE-WV-L278-004,CLIENT,Oakwood
GM1B9L5Z,LBE-WV-L278-005,CLIENT,Oakwood
GM1B9L64,LBE-WV-L296-201,ADMIN,Palmers Green - Ground Floor
GM1B9L65,LBE-WV-L296-001,CLIENT,Palmers Green - Ground Floor
GM1B9L66,LBE-WV-L296-002,CLIENT,Palmers Green - Ground Floor
GM1B9L67,LBE-WV-L296-003,CLIENT,Palmers Green - Ground Floor
GM1B9L68,LBE-WV-L296-004,CLIENT,Palmers Green - Ground Floor
GM1B9L6A,LBE-WV-L297-201,ADMIN,Palmers Green -  First Floor
GM1B9L6B,LBE-WV-L297-001,CLIENT,Palmers Green -  First Floor
GM1B9L6C,LBE-WV-L297-002,CLIENT,Palmers Green -  First Floor
GM1B9L6D,LBE-WV-L297-003,CLIENT,Palmers Green -  First Floor
GM1B9L6E,LBE-WV-L297-004,CLIENT,Palmers Green -  First Floor
GM1B9L6G,LBE-WV-L297-005,CLIENT,Palmers Green -  First Floor
GM1B9L6J,LBE-WV-L297-006,CLIENT,Palmers Green -  First Floor
GM1B9L6K,LBE-WV-L297-007,CLIENT,Palmers Green -  First Floor
GM1B9L6Z,LBE-WV-L297-008,CLIENT,Palmers Green -  First Floor
GM1B9L72,LBE-WV-L297-009,CLIENT,Palmers Green -  First Floor
GM1B9L74,LBE-WV-L297-010,CLIENT,Palmers Green -  First Floor
GM1B9L75,LBE-WV-L297-011,CLIENT,Palmers Green -  First Floor
GM1B9L7D,LBE-WV-L297-012,CLIENT,Palmers Green -  First Floor
GM1B9L7N,LBE-WV-L272-201,ADMIN,RIDGE AVENUE
GM1B9L7P,LBE-WV-L272-001,CLIENT,RIDGE AVENUE
GM1B9L7Q,LBE-WV-L272-002,CLIENT,RIDGE AVENUE
GM1B9L7W,LBE-WV-L272-003,CLIENT,RIDGE AVENUE
GM1B9L7Y,LBE-WV-L272-004,CLIENT,RIDGE AVENUE
GM1B9L7Z,LBE-WV-L272-005,CLIENT,RIDGE AVENUE
GM1B9L80,LBE-WV-L272-006,CLIENT,RIDGE AVENUE
GM1B9L81,LBE-WV-L272-007,CLIENT,RIDGE AVENUE
GM1B9L82,LBE-WV-L272-008,CLIENT,RIDGE AVENUE
GM1B9L83,LBE-WV-L272-009,CLIENT,RIDGE AVENUE
GM1B9L84,LBE-WV-L272-010,CLIENT,RIDGE AVENUE
GM1B9L85,LBE-WV-L272-011,CLIENT,RIDGE AVENUE
GM1B9L86,LBE-WV-L282-201,ADMIN,FORE STREET
GM1B9L87,LBE-WV-L282-001,CLIENT,FORE STREET
GM1B9L88,LBE-WV-L282-002,CLIENT,FORE STREET
GM1B9L89,LBE-WV-L282-003,CLIENT,FORE STREET
GM1B9L8A,LBE-WV-L282-004,CLIENT,FORE STREET
GM1B9L8C,LBE-WV-L282-005,CLIENT,FORE STREET
GM1B9L8D,LBE-WV-L282-006,CLIENT,FORE STREET
GM1B9L8E,LBE-WV-L282-007,CLIENT,FORE STREET
GM1B9L8F,LBE-WV-L282-008,CLIENT,FORE STREET
GM1B9L8G,LBE-WV-L282-009,CLIENT,FORE STREET
GM1B9L8H,LBE-WV-L282-010,CLIENT,FORE STREET
GM1B9L8K,LBE-WV-L282-011,CLIENT,FORE STREET
GM1B9L8L,LBE-WV-L282-012,CLIENT,FORE STREET
GM1B9L8M,LBE-WV-L287-201,ADMIN,Millfield
GM1B9L8N,LBE-WV-L287-001,CLIENT,Millfield
GM1B9L8P,LBE-WV-L287-002,CLIENT,Millfield
GM1B9L8R,LBE-WV-L273-201,ADMIN,Ordnance Unit Centre - Ground Floor
GM1B9L8Y,LBE-WV-L273-001,CLIENT,Ordnance Unit Centre - Ground Floor
GM1B9L8Z,LBE-WV-L273-002,CLIENT,Ordnance Unit Centre - Ground Floor
GM1B9L90,LBE-WV-L273-003,CLIENT,Ordnance Unit Centre - Ground Floor
GM1B9L91,LBE-WV-L273-004,CLIENT,Ordnance Unit Centre - Ground Floor
GM1B9L92,LBE-WV-L273-005,CLIENT,Ordnance Unit Centre - Ground Floor
GM1B9L93,LBE-WV-L273-006,CLIENT,Ordnance Unit Centre - Ground Floor
GM1B9L94,LBE-WV-L273-007,CLIENT,Ordnance Unit Centre - Ground Floor
GM1B9L95,LBE-WV-L273-008,CLIENT,Ordnance Unit Centre - Ground Floor
GM1B9L96,LBE-WV-L273-009,CLIENT,Ordnance Unit Centre - Ground Floor
GM1B9L9C,LBE-WV-L273-010,CLIENT,Ordnance Unit Centre - Ground Floor
GM1B9L9N,LBE-WV-L292-201,ADMIN,Ordnance Unit Centre - First Floor
GM1B9L9P,LBE-WV-L292-001,CLIENT,Ordnance Unit Centre - First Floor
GM1B9L9Q,LBE-WV-L292-002,CLIENT,Ordnance Unit Centre - First Floor
GM1B9L9V,LBE-WV-L292-003,CLIENT,Ordnance Unit Centre - First Floor
GM1B9L9W,LBE-WV-L292-004,CLIENT,Ordnance Unit Centre - First Floor
GM1B9LA0,LBE-WV-L292-005,CLIENT,Ordnance Unit Centre - First Floor
GM1B9LA2,LBE-WV-L292-006,CLIENT,Ordnance Unit Centre - First Floor
GM1AZHY3,LBE-WV-L292-007,CLIENT,Ordnance Unit Centre - First Floor
GM1B9L7X,LBE-WV-L292-008,CLIENT,Ordnance Unit Centre - First Floor
GM1B9L7L,LBE-WV-L274-201,ADMIN,Ponders End
GM1AZHY4,LBE-WV-L274-001,CLIENT,Ponders End
GM1B9L5Q,LBE-WV-L274-002,CLIENT,Ponders End
GM1B9L5T,LBE-WV-L274-003,CLIENT,Ponders End
GM1B9L69,LBE-WV-L274-004,CLIENT,Ponders End
GM1B9L6M,LBE-WV-L274-005,CLIENT,Ponders End
GM1B9L6N,LBE-WV-L274-006,CLIENT,Ponders End
GM1B9L6Q,LBE-WV-L274-007,CLIENT,Ponders End
GM1B9L6R,LBE-WV-L294-201,ADMIN,Edmonton Library - Ground Floor
GM1B9L6T,LBE-WV-L294-001,CLIENT,Edmonton Library - Ground Floor
GM1B9L6W,LBE-WV-L294-002,CLIENT,Edmonton Library - Ground Floor
GM1B9L6X,LBE-WV-L294-003,CLIENT,Edmonton Library - Ground Floor
GM1B9L73,LBE-WV-L294-004,CLIENT,Edmonton Library - Ground Floor
GM1B9L76,LBE-WV-L294-005,CLIENT,Edmonton Library - Ground Floor
GM1B9L79,LBE-WV-L294-006,CLIENT,Edmonton Library - Ground Floor
GM1B9L7V,LBE-WV-L294-007,CLIENT,Edmonton Library - Ground Floor
GM1B9L8T,LBE-WV-L294-008,CLIENT,Edmonton Library - Ground Floor
GM1B9L9F,LBE-WV-L294-009,CLIENT,Edmonton Library - Ground Floor
GM1B9L9G,LBE-WV-L294-010,CLIENT,Edmonton Library - Ground Floor
GM1B9L9H,LBE-WV-L294-011,CLIENT,Edmonton Library - Ground Floor
GM1B9L9J,LBE-WV-L294-012,CLIENT,Edmonton Library - Ground Floor
GM1B9L9L,LBE-WV-L294-013,CLIENT,Edmonton Library - Ground Floor
GM1B9L9Y,LBE-WV-L294-014,CLIENT,Edmonton Library - Ground Floor
GM1B9LA1,LBE-WV-L294-015,CLIENT,Edmonton Library - Ground Floor
GM1EWS76,LBE-WV-L294-016,CLIENT,Edmonton Library - Ground Floor
GM1EWS72,LBE-WV-L294-017,CLIENT,Edmonton Library - Ground Floor
GM1EWSCF,LBE-WV-L294-018,CLIENT,Edmonton Library - Ground Floor
GM1EWSCS,LBE-WV-L294-019,CLIENT,Edmonton Library - Ground Floor
GM1EWS7L,LBE-WV-L294-020,CLIENT,Edmonton Library - Ground Floor
GM1EWS74,LBE-WV-L294-021,CLIENT,Edmonton Library - Ground Floor
GM1EWS7Y,LBE-WV-L294-022,CLIENT,Edmonton Library - Ground Floor
GM1EWS63,LBE-WV-L294-023,CLIENT,Edmonton Library - Ground Floor
GM1EWS8F,LBE-WV-L294-024,CLIENT,Edmonton Library - Ground Floor
GM1EWS6T,LBE-WV-L294-025,CLIENT,Edmonton Library - Ground Floor
GM1EWSCE,LBE-WV-L294-026,CLIENT,Edmonton Library - Ground Floor
GM1EWS7K,LBE-WV-L294-027,CLIENT,Edmonton Library - Ground Floor
GM1EWS6M,LBE-WV-L294-028,CLIENT,Edmonton Library - Ground Floor
GM1EWSC1,LBE-WV-L294-029,CLIENT,Edmonton Library - Ground Floor
GM1EWS6F,LBE-WV-L294-030,CLIENT,Edmonton Library - Ground Floor
GM1EWS8R,LBE-WV-L294-031,CLIENT,Edmonton Library - Ground Floor
GM1EWS93,LBE-WV-L294-032,CLIENT,Edmonton Library - Ground Floor
GM1EWS82,LBE-WV-L294-033,CLIENT,Edmonton Library - Ground Floor
GM1EWS8L,LBE-WV-L294-034,CLIENT,Edmonton Library - Ground Floor
GM1EWS5H,LBE-WV-L294-035,CLIENT,Edmonton Library - Ground Floor
GM1EWS5Z,LBE-WV-L294-036,CLIENT,Edmonton Library - Ground Floor
GM1EWS7H,TBC,SKIP,Edmonton Library - Ground Floor - Community Hub
GM1EWS7B,TBC,SKIP,Edmonton Library - Ground Floor - Community Hub
GM1EWS6H,TBC,SKIP,Edmonton Library - Ground Floor - Community Hub
GM1EWS99,TBC,SKIP,Edmonton Library - Ground Floor - Community Hub
GM1B9L7K,TBC,SKIP,Edmonton Library - Ground Floor - Community Hub
GM1B9L7J,TBC,SKIP,Edmonton Library - Ground Floor - Community Hub
GM1B9L6L,LBE-WV-L295-201,ADMIN,Edmonton Library - First Floor
GM1B9L7E,LBE-WV-L295-001,CLIENT,Edmonton Library - First Floor
GM1B9L7H,LBE-WV-L295-002,CLIENT,Edmonton Library - First Floor
TBC,LBE-WV-L295-003,SKIP,Edmonton Library - First Floor
TBC,LBE-WV-L295-004,SKIP,Edmonton Library - First Floor
TBC,LBE-WV-L295-005,SKIP,Edmonton Library - First Floor
TBC,LBE-WV-L295-006,SKIP,Edmonton Library - First Floor
TBC,LBE-WV-L295-007,SKIP,Edmonton Library - First Floor
TBC,LBE-WV-L295-008,SKIP,Edmonton Library - First Floor
TBC,LBE-WV-L295-009,SKIP,Edmonton Library - First Floor
TBC,LBE-WV-L295-010,SKIP,Edmonton Library - First Floor
TBC,LBE-WV-L295-011,SKIP,Edmonton Library - First Floor
TBC,LBE-WV-L295-012,SKIP,Edmonton Library - First Floor
TBC,LBE-WV-L295-013,SKIP,Edmonton Library - First Floor
TBC,LBE-WV-L295-014,SKIP,Edmonton Library - First Floor
TBC,LBE-WV-L295-015,SKIP,Edmonton Library - First Floor
TBC,LBE-WV-L295-016,SKIP,Edmonton Library - First Floor
TBC,LBE-WV-L295-017,SKIP,Edmonton Library - First Floor
TBC,LBE-WV-L295-018,SKIP,Edmonton Library - First Floor
TBC,LBE-WV-L295-019,SKIP,Edmonton Library - First Floor
TBC,LBE-WV-L295-020,SKIP,Edmonton Library - First Floor
TBC,LBE-WV-L295-021,SKIP,Edmonton Library - First Floor
TBC,LBE-WV-L295-022,SKIP,Edmonton Library - First Floor
TBC,LBE-WV-L295-023,SKIP,Edmonton Library - First Floor
TBC,LBE-WV-L295-024,SKIP,Edmonton Library - First Floor
TBC,LBE-WV-L295-025,SKIP,Edmonton Library - First Floor
TBC,LBE-WV-L295-026,SKIP,Edmonton Library - First Floor
TBC,LBE-WV-L295-027,SKIP,Edmonton Library - First Floor
TBC,LBE-WV-L295-028,SKIP,Edmonton Library - First Floor
'@ | ConvertFrom-Csv

Write-Log '===== Script started ====='

# 1. Serial number
$Serial = (Get-CimInstance -ClassName Win32_BIOS).SerialNumber
if ($Serial) { $Serial = $Serial.Trim() }
Write-Log "BIOS serial number: $Serial"

$Device = $DeviceTable | Where-Object { $_.Serial -eq $Serial } | Select-Object -First 1
if (-not $Device) {
    Write-Log "Serial $Serial is not in the table. No changes made."
    exit 0
}
if ($Device.Role -eq 'SKIP' -or $Device.Name -eq 'TBC') {
    Write-Log "Serial $Serial has no computer name yet (TBC). No changes made."
    exit 0
}

$NewName = $Device.Name.Trim()
Write-Log "Matched: Name=$NewName Role=$($Device.Role) Location=$($Device.Location)"

# 2. iCAM registry settings (library customer PCs only)
if ($Device.Role -eq 'CLIENT') {
    if ($NewName -match '^(LBE-WV-L\d{3})-') {
        $AdminHost = "$($Matches[1])-201"
        Write-Log "Site admin PC: $AdminHost"

        $Keys = @(
            @{ Path = 'HKLM:\SOFTWARE\Insight Media\Cafe Client';  Name = 'ServerAddress' },
            @{ Path = 'HKLM:\SOFTWARE\Insight Media\Print Client'; Name = 'Server IP' }
        )
        foreach ($k in $Keys) {
            try {
                if (-not (Test-Path $k.Path)) { New-Item -Path $k.Path -Force | Out-Null }
                New-ItemProperty -Path $k.Path -Name $k.Name -Value $AdminHost -PropertyType String -Force | Out-Null
                Write-Log "Set $($k.Path)\$($k.Name) = $AdminHost"
            } catch {
                Write-Log "ERROR setting $($k.Path)\$($k.Name): $($_.Exception.Message)"
            }
        }
    } else {
        Write-Log "Could not work out the site code from $NewName. Registry not changed."
    }
} else {
    Write-Log "Role is $($Device.Role): iCAM registry settings not applied."
}

# 3. Rename
if ($env:COMPUTERNAME -ieq $NewName) {
    Write-Log "Computer is already named $NewName. No rename needed."
    Write-Log '===== Script finished ====='
    exit 0
}
try {
    Rename-Computer -NewName $NewName -Force -ErrorAction Stop
    Write-Log "Renamed $env:COMPUTERNAME to $NewName. Restarting in 30 seconds."
} catch {
    Write-Log "ERROR renaming computer: $($_.Exception.Message)"
    exit 1
}

Write-Log '===== Script finished ====='
shutdown.exe /r /t 30 /c "Applying new computer name $NewName"
exit 0
