param (
    [string]$inncodes,
    [string]$ssofile,
    [string]$savedirectory
)

#IMPORT SSO LOGON
$sso_login = ConvertFrom-Json -InputObject (Get-Content -Path $ssofile -Raw)

write-host "$inncodes"

$inncodeArray = @()
$inncodeArray = $inncodes -split "," | ForEach-Object { $_.Trim() }

foreach ($inncode in $inncodeArray){

#ESTABLISH CONNECTION DETAILS
$session = New-Object Microsoft.PowerShell.Commands.WebRequestSession
$session.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36 Edg/131.0.0.0"
$user_agent = $session.UserAgent
$user_id = $sso_login.user.username
$hk_info = $user_agent + ":ADMIN:0.9.1:" + $user_id + ":0000000000000::::"
$hk_info = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($hk_info))

$headers = @{
  "Accept"             = "application/json, text/plain, */*"
  "Accept-Encoding"    = "identity"
  "Accept-Language"    = "en-US,en;q=0.9"
  "Authorization"      = "Bearer $($sso_login.token)"
  "Origin"             = "https://login.pep.hilton.com"
  "Referer"            = "https://login.pep.hilton.com/"
  "Sec-Fetch-Dest"     = "empty"
  "Sec-Fetch-Mode"     = "cors"
  "Sec-Fetch-Site"     = "same-site"
  "hk-app-id"          = "ADMIN"
  "hk-app-version"     = "4.5.5.2.53.7"
  "hk-info"            = "$hk_info"
  "sec-ch-ua"          = "`"Microsoft Edge`";v=`"131`", `"Chromium`";v=`"131`", `"Not_A Brand`";v=`"24`""
  "sec-ch-ua-mobile"   = "?0"
  "sec-ch-ua-platform" = "`"Windows`""
}

$currentProperty = $sso_login.properties | Where-Object { $_.code -eq $inncode }

Write-Host "`nCurrent Property: $inncode"

$propertyID = $currentProperty.id

$session = New-Object Microsoft.PowerShell.Commands.WebRequestSession
$session.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/133.0.0.0 Safari/537.36 Edg/133.0.0.0"
$pbxConfigResponse = Invoke-WebRequest -UseBasicParsing -Uri "$($currentProperty.region.url)hk-property-interfaces/hotelbrand/properties/$($currentProperty.id)/system-config/pbx" -WebSession $session -Headers $headers
$pbxConfig = $pbxConfigResponse.Content | ConvertFrom-Json

$createHSKPBody = "{`"stream_id`":`"$($pbxConfig.stream_id)`",`"property_id`":`"$propertyID`",`"enabled`":$($pbxConfig.enabled),`"room_alias_type`":`"$($pbxConfig.room_alias_type)`",`"aws_thing_name`":`"pep-prod-$($inncode)-1`",`"room_alias_mappings`":[],`"housekeeping_status_mappings`":[{`"status_code`":`"1`",`"occupied_status`":`"OCCUPIED`",`"clean_status`":`"CLEAN`"},{`"status_code`":`"2`",`"occupied_status`":`"OCCUPIED`",`"clean_status`":`"DIRTY`"},{`"status_code`":`"3`",`"occupied_status`":`"VACANT`",`"clean_status`":`"CLEAN`"},{`"status_code`":`"4`",`"occupied_status`":`"VACANT`",`"clean_status`":`"DIRTY`"},{`"status_code`":`"5`",`"occupied_status`":`"OCCUPIED`",`"clean_status`":`"READY`"},{`"status_code`":`"7`",`"occupied_status`":`"VACANT`",`"clean_status`":`"READY`"}]}"
Invoke-WebRequest -UseBasicParsing -Uri "$($currentProperty.region.url)hk-property-interfaces/hotelbrand/properties/$($currentProperty.id)/system-config/pbx" -Method "PUT" -WebSession $session -Headers $headers -ContentType "application/json;charset=UTF-8" -Body $createHSKPBody
}

# Add your logic to handle the SSO file and other operations here
Write-Output "Script execution completed."

Read-Host "Press Enter to close the window"
exit

