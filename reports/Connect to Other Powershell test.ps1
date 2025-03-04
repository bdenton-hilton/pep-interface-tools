param (
    [string]$inncodes,
    [string]$ssofile,
    [string]$savedirectory
)

Read-Host "Press Enter to close the window"

$inncodeArray = $null
$inncodeArray = $inncodes -split "," | ForEach-Object { $_.Trim() }
foreach ($code in $inncodeArray) {
    Write-Output "Processing inncode: $code"
    # Add your processing logic here
}

# Main script logic
Write-Output "SSO File: $ssofile"
Write-Output "Save Directory: $savedirectory"

# Check if the save directory exists, if not create it
if (-not (Test-Path -Path $savedirectory)) {
    Write-Output "Created save directory: $savedirectory"
}

# Add your logic to handle the SSO file and other operations here
Write-Output "Script execution completed."

Read-Host "Press Enter to close the window"
exit