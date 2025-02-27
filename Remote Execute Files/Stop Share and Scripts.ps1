# This script will stop the sharing of the scripts and userdata folders and rename a file

# Show the shared folders before sharing is stopped
Get-SmbShare

# Stops scripts sharing
Remove-SmbShare -Name "scripts" -Force

# Stops userdata sharing
Remove-SmbShare -Name "userdata" -Force

# Show the shared folders after sharing has been stopped
Get-SmbShare

# Rename the local.dli file with extension .PEP
Rename-Item -Path "D:\Install\support\logon\scripts\local.dli" -NewName "local.dli.PEP"

Write-Host
Pause
