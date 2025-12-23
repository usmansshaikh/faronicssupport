# Installs Microsoft 365 Apps (French) using Office Deployment Tool (ODT)
# Works in SYSTEM context (Deploy/Intune/RMM). Downloads everything from Microsoft CDN.

$ErrorActionPreference = "Stop"

$WorkDir = Join-Path $env:ProgramData "ODT-M365-frFR"
New-Item -ItemType Directory -Path $WorkDir -Force | Out-Null

# 1) Download ODT (Microsoft official)
$OdtBootstrap = Join-Path $WorkDir "odt-setup.exe"
Invoke-WebRequest -Uri "https://officecdn.microsoft.com/pr/wsus/setup.exe" -OutFile $OdtBootstrap

# 2) Extract ODT
Start-Process -FilePath $OdtBootstrap -ArgumentList "/extract:`"$WorkDir`" /quiet" -Wait

$SetupExe = Join-Path $WorkDir "setup.exe"
if (-not (Test-Path $SetupExe)) {
    throw "ODT extraction failed: setup.exe not found in $WorkDir"
}

# 3) Create config (French)
# Notes:
# - OfficeClientEdition="64" -> 64-bit Office
# - Channel="Current" -> generally "latest"; change if you need a specific channel
# - Product ID "O365ProPlusRetail" -> Microsoft 365 Apps for enterprise
$ConfigXml = Join-Path $WorkDir "configuration.xml"

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
</Configuration>
"@ | Set-Content -Path $ConfigXml -Encoding UTF8

# 4) Install (downloads from Microsoft CDN automatically)
Start-Process -FilePath $SetupExe -ArgumentList "/configure `"$ConfigXml`"" -Wait

Write-Host "SUCCESS: Microsoft 365 Apps installed/updated in French (fr-FR)."
exit 0
