# Get and format the property code
$hostName = (hostname)
$propertyCode = $hostName.Substring(0, 5)

# Get and format the username
$username = (whoami).Split('-')[1]

# Get and format the current date
$currentDate = Get-Date -Format "MMddyyyy"

# Log file path
$logFilePath = "D:\datatransfer_all-$currentDate`_$username.txt"

# Inform user and run robocopy
Write-Output "Robocopy will now run. Property:$propertyCode, User:$username, Date:$currentDate"

# Run our existing robocopy code, but with dynamic items (property code, username, date)
Start-Process robocopy -ArgumentList "G:\", "\\$propertyCode`leg1\userdata", "-xf *.iso *.wim /E /s /w:5 /R:3 /xd `$RECYCLE.BIN 'System Volume Information' /MAXAGE:730 /XN /log:`"$logFilePath`"" -NoNewWindow

# Tail the log file while robocopy is running
Get-Content -Path $logFilePath -Wait -Tail 10
