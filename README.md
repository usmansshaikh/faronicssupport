# Library PC Lockdown

Designed for a Windows public-library workstation with a local standard user named `Public`.

## Files

- `Deploy-LibraryLockdown.ps1` - applies Public-user restrictions.
- `Audit-LibraryLockdown.ps1` - checks the main restrictions.
- `Remove-LibraryLockdown.ps1` - removes restrictions added by Deploy.

## Prerequisites

1. `Public` must exist as a local standard user.
2. `Public` must NOT be a member of Administrators.
3. Sign into `Public` at least once so Windows creates its profile.
4. Test on one workstation before mass deployment.

## Deep Freeze Cloud

After uploading these files to a public GitHub repository, the Deploy script can be downloaded and executed remotely as Administrator/SYSTEM.

Example, replace OWNER/REPO as required:

    powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "irm 'https://raw.githubusercontent.com/OWNER/REPO/main/library-lockdown/Deploy-LibraryLockdown.ps1' | iex"

The scripts target only the specified Public user's registry hive. Administrator accounts are not intentionally restricted.

## Secure Boot Check

The `scripts/SecureBootCheck.ps1` script checks whether the Windows UEFI CA 2023 certificate is present in the Secure Boot db. This is a read-only diagnostic that can be executed remotely as SYSTEM via Deep Freeze Cloud:

    powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "iex (irm 'https://raw.githubusercontent.com/usmansshaikh/faronicssupport/main/scripts/SecureBootCheck.ps1')"

The script logs results to `C:\ProgramData\SecureBootCheck.txt` and exits with code 0 if the certificate is present, or 1 if missing, unsupported, or an error occurred (e.g., Legacy BIOS).

## Set-NeverSleep

The `scripts/Set-NeverSleep.ps1` script sets Sleep, Hibernate, and Hybrid Sleep to Never on all power plans (AC + DC) and disables hibernation completely (removes hiberfil.sys, disables Fast Startup). This can be executed remotely as SYSTEM via Deep Freeze Cloud:

    powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "iex (irm 'https://raw.githubusercontent.com/usmansshaikh/faronicssupport/main/scripts/Set-NeverSleep.ps1')"

**Note:** Thaw PCs before running this script, as it modifies system power settings and requires write access.

The script logs activity to `C:\Windows\Temp\Set-NeverSleep.log`.

## Auto-logon

Auto-logon is deliberately kept separate because a public GitHub repository must not contain the Public account password. Use Microsoft Sysinternals Autologon or another centrally managed credential mechanism.

## Recovery

Run:

    .\Remove-LibraryLockdown.ps1 -User Public

Then sign out/restart Windows.
