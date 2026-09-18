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

## Auto-logon

Auto-logon is deliberately kept separate because a public GitHub repository must not contain the Public account password. Use Microsoft Sysinternals Autologon or another centrally managed credential mechanism.

## Recovery

Run:

    .\Remove-LibraryLockdown.ps1 -User Public

Then sign out/restart Windows.
