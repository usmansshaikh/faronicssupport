# SecureBootCheck.ps1
# Purpose: Check whether the Windows UEFI CA 2023 certificate is present in the Secure Boot db.
# This script is read-only; it does not modify any system settings.
# Designed to be executed remotely as SYSTEM via Deep Freeze Cloud using: iex (irm <raw URL>)
#
# Exit codes:
#   0 = Windows UEFI CA 2023 certificate is present
#   1 = Certificate missing, error occurred, or Secure Boot unsupported (Legacy BIOS)
#
# Log file: C:\ProgramData\SecureBootCheck.txt

$log = "C:\ProgramData\SecureBootCheck.txt"
try {
    $sb = Confirm-SecureBootUEFI
    $ca = [System.Text.Encoding]::ASCII.GetString((Get-SecureBootUEFI db).bytes) -match 'Windows UEFI CA 2023'
    $result = "SecureBoot=$sb; UEFI_CA_2023=$ca"
} catch {
    $result = "Error: $($_.Exception.Message) (likely Legacy BIOS or Secure Boot unsupported)"
}
"$env:COMPUTERNAME; $(Get-Date -Format s); $result" | Out-File $log -Encoding ascii
Write-Output $result
exit $(if ($ca) {0} else {1})
