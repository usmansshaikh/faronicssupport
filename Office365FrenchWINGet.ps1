$ErrorActionPreference = 'Stop'

Write-Host "=== STEP 1: Check / Install WinGet ==="

$WingetPath = "$env:LOCALAPPDATA\Microsoft\WindowsApps\winget.exe"

if (-not (Test-Path $WingetPath)) {

    Write-Host "WinGet not found. Installing App Installer (WinGet)..."

    $Temp = "$env:TEMP\WinGet"
    New-Item -ItemType Directory -Path $Temp -Force | Out-Null

    # Download dependencies
    Invoke-WebRequest "https://aka.ms/Microsoft.VCLibs.x64.14.00.Desktop.appx" `
        -OutFile "$Temp\VCLibs.appx"

    Invoke-WebRequest "https://aka.ms/getwinget" `
        -OutFile "$Temp\AppInstaller.msixbundle"

    # Install dependencies
    Add-AppxPackage "$Temp\VCLibs.appx"

    # Install App Installer (includes WinGet)
    Add-AppxPackage "$Temp\AppInstaller.msixbundle"

    Start-Sleep -Seconds 5
}

if (-not (Get-Command winget -ErrorAction SilentlyContinue)) {
    Write-Error "WinGet installation failed."
    exit 1
}

Write-Host "WinGet is installed."

# ------------------------------------------------------------

Write-Host "=== STEP 2: Install / Upgrade Microsoft 365 Apps (French) ==="

$OfficeId = "Microsoft.Office"
$Locale   = "fr-FR"

# Try upgrade first
winget upgrade --id $OfficeId `
               --exact `
               --silent `
               --locale $Locale `
               --accept-package-agreements `
               --accept-source-agreements

# If upgrade failed or Office not present → install
if ($LASTEXITCODE -ne 0) {
    Write-Host "Office not found or upgrade failed. Installing..."
    
    winget install --id $OfficeId `
                   --exact `
                   --silent `
                   --locale $Locale `
                   --accept-package-agreements `
                   --accept-source-agreements
}

Write-Host "=== Microsoft 365 Apps deployment completed ==="