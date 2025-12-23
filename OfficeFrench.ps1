# Microsoft 365 Apps (French) install/update via ODT - NO EXTRACT
# Works under SYSTEM (Faronics Deploy / Intune / RMM)

$ErrorActionPreference = "Stop"

$WorkDir  = Join-Path $env:ProgramData "ODT-M365-frFR"
$SetupExe = Join-Path $WorkDir "setup.exe"
$Config   = Join-Path $WorkDir "configuration.xml"

New-Item -ItemType Directory -Path $WorkDir -Force | Out-Null

# Download ODT (this file IS the ODT executable)
Invoke-WebRequest -Uri "https://officecdn.microsoft.com/pr/wsus/setup.exe" -OutFile $SetupExe

# Verify file exists
if (-not (Test-Path $SetupExe)) {
    throw "ODT download failed: setup.exe not found in $WorkDir"
}

# Verify it's a real EXE (MZ header). If not, proxy/captive portal likely returned HTML.
$bytes = [System.IO.File]::ReadAllBytes($SetupExe)
if ($bytes.Length -lt 2 -or $bytes[0] -ne 0x4D -or $bytes[1] -ne 0x5A) {
    $head = [System.Text.Encoding]::UTF8.GetString($bytes, 0, [Math]::Min($bytes.Length, 300))
    throw "Downloaded setup.exe is not a valid EXE (MZ header missing). Likely blocked/rewritten download. First bytes:`n$head"
}

# Create ODT config (French fr-FR, 64-bit, latest Current channel)
@"
<Configuration>
  <Add OfficeClientEdition="64" Channel="Current">
    <Product ID="O365ProPlusRetail">
      <Language ID="fr-fr" />
    </Product>
  </Add>

  <Display Level="None" AcceptEULA="TRUE" />
  <Property Name="FORCEAPPSHUTDOWN" Value="TRUE" />
  <Updates Enabled="TRUE" />
  <Logging Level="Standard" Path="$WorkDir" />
</Configuration>
"@ | Set-Content -Path $Config -Encoding UTF8

# Install/update Office (downloads content from Microsoft CDN)
$proc = Start-Process -FilePath $SetupExe -ArgumentList "/configure `"$Config`"" -Wait -PassThru
if ($proc.ExitCode -ne 0) {
    throw "ODT failed with ExitCode $($proc.ExitCode). Check logs in $WorkDir"
}

Write-Host "SUCCESS: Microsoft 365 Apps installed/updated in French. Logs: $WorkDir"