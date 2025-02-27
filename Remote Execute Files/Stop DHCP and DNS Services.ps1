Write-Output "Current DHCP and DNS service statuses..."

$dhcpServerService = Get-Service DHCPServer
Write-Output $dhcpServerService

Write-Output "Stopping DHCP and DNS services..."

Stop-Service -Name "DHCPServer" -Verbose
Set-Service -Name "DHCPServer" -StartupType Disabled -Verbose

Stop-Service -Name "DNS" -Verbose
Set-Service -Name "DNS" -StartupType Disabled -Verbose

Write-Output "New DHCP and DNS service statuses..."

$dhcpServerService = Get-Service DHCPServer
Write-Output $dhcpServerService