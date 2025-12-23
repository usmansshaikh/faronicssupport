$pkg = "Kaspersky.ShellEx22"

# 1) Remove the installed package for any existing users
Get-AppxPackage -AllUsers | Where-Object {$_.Name -like $pkg} | ForEach-Object {
  Write-Host "Removing installed package:" $_.PackageFullName
  Remove-AppxPackage -Package $_.PackageFullName -AllUsers
}

# 2) Remove it from the image provisioning (so it won't come back for new users)
Get-AppxProvisionedPackage -Online | Where-Object {$_.DisplayName -like $pkg} | ForEach-Object {
  Write-Host "Removing provisioned package:" $_.PackageName
  Remove-AppxProvisionedPackage -Online -PackageName $_.PackageName
}