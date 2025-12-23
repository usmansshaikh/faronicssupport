# Installs Microsoft 365 Apps (French) using Office Deployment Tool (ODT)
# Downloads ODT from Microsoft CDN and installs silently.
# Works under SYSTEM (Deploy/Intune/RMM).

$ErrorActionPreference = "Stop"

$WorkDir = Join-Path $env:ProgramData "ODT-M365-frFR"
New-Item -ItemType Directory -Path $WorkDir -Force | Out-Null

$SetupExe   = Join-Path $WorkDir "setup.exe"
$ConfigXml  = Join-Path $WorkDir "configuration.xml"
$LogPath    = Join-Path $WorkDir "odt-install.log"

# 1) Download ODT setup.exe (this IS the ODT executable)
Invoke-WebRequest -Uri "https://officecdn.microsoft.com/pr/wsus/setup.exe" -OutFile $SetupExe

# 2) Validate we actually got an EXE (not HTML from proxy/login page)
$bytes = [System.IO.File]::ReadAllBytes($SetupExe)
if ($bytes.Length -lt 2 -or $bytes[0] -ne 0x4D -or $bytes[1] -ne 0x5A) {
    $head = [System.Text.Encoding]::UTF8.GetString($bytes, 0, [Math]::Min($bytes.Length, 300))
    throw "Downloaded setup.exe is not a valid Windows executable (MZ header missing). First bytes:`n$head"
}

# 3) Create ODT configuration (French)
@"
<Configuration>
  <Add OfficeClientEdition="64" Channel="Current">
    <Product ID="O365ProPlusRetail">
      <Language ID="fr-fr" />
    </Product>
  </Add>

  <Display Level="None" AcceptEULA="TRUE" />
  <Property Name="FORCEAPPSHUTDOWN" Value="TRUE" />
  <Property Name="AUTOACTIVATE" Value="1" />
  <Updates Enabled="TRUE" />
  <Logging Level="Standard" Path="$WorkDir" />
</Configuration>
"@ | Set-Content -Path $ConfigXml -Encoding UTF8

# 4) Install (downloads Office from Microsoft CDN automatically)
$proc = Start-Process -FilePath $SetupExe -ArgumentList "/configure `"$ConfigXml`"" -Wait -PassThru

if ($proc.ExitCode -ne 0) {
    throw "ODT setup.exe returned ExitCode $($proc.ExitCode). Check logs in: $WorkDir"
}

"SUCCESS: Microsoft 365 Apps installed/updated in French (fr-FR)." | Out-File -FilePath $LogPath -Append -Encoding UTF8
Write-Host "SUCCESS: Microsoft 365 Apps installed/updated in French (fr-FR). Logs: $WorkDir"
exit 0
