param (
    [string]$inncodes,
    [string]$ssofile,
    [string]$savedirectory,
    [string]$tempFilePath
)

$globalSettings = Import-Clixml -Path $tempFilePath
Remove-Item -Path $tempFilePath -Force

function encyptedPlaintextPasswordToCredentials {
    param (
        $username,
        $encryptedpassword
    ) 
    return New-Object PSCredential ($username, (ConvertTo-SecureString $encryptedpassword -ErrorAction SilentlyContinue))
}

if ($env:globalSettings.Credentials.'NA-ADM Password'.defaultvalue -ne "password")
{ $credential = encyptedPlaintextPasswordToCredentials -username $($globalSettings.Credentials.'NA-ADM Username'.defaultvalue) -encryptedpassword $($globalSettings.Credentials.'NA-ADM Password'.defaultvalue) }
else { $credential = Get-Credential }

#IMPORT SSO LOGON
$sso_login = ConvertFrom-Json -InputObject (Get-Content -Path $ssofile -Raw)

write-host "$inncodes"

$inncodeArray = $inncodes -split "," | ForEach-Object { $_.Trim() }

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

$reportResult = "Selected Inn Codes:`n"
$reportResult += $inncodes
foreach ($inncode in $inncodeArray) {
    $reportResult += "`n" + "-------- $inncode --------" + "`n"

    $currentProperty = $sso_login.properties | Where-Object { $_.code -eq $inncode }

    Write-Host "`nCurrent Property: $inncode"

    $nameAndHost = "SERVER.NA.HHCPR.HILTON.COM"
    $computer = $inncode.ToUpper() + $nameAndHost

    $filteredMaidCodes = $null
    $selectMaidCodes = @"
SELECT
    X.input_code,
    F.code_desc
FROM
    [hpms3].[dbo].[LIMS_MAID_STATUS_CODES_FIXED] F
JOIN
    [hpms3].[dbo].[LIMS_MAID_STATUS_CODES_XLTED] X
ON
    X.translated_code = F.code
"@

    $targetScriptBlock = { invoke-sqlcmd -database "hpms3" -query $using:selectMaidCodes }


    Write-Host "`nConnecting to $computer"

    $commandOutput = $null
    $commandOutput = Invoke-Command -ComputerName $computer -Credential $credential -ScriptBlock $targetScriptBlock


    #############################################################################

    $filteredMaidCodes = $($commandOutput | Where-Object { $_.input_code -notmatch '[a-zA-Z]' } | Where-Object { $_.code_desc -notlike '*Attendant*' } | Select-Object input_code, code_desc)

    if ($filteredMaidCodes) {
        $maidCodeStrings = @()
        $input_code = $null

        Write-Host "Housekeeping Codes Found:"
        Write-Host "`n"
        $filteredMaidCodes | Format-Table

        foreach ($entry in $filteredMaidCodes) {
            $input_code = $entry.input_code
            $statuses = $statuses -replace '(^\S+\s+\S+)\s.*', '$1' #regex to remove anything after the second word in the status string.
            $statuses = $entry.code_desc.ToUpper() -split ' '
            $occupied_status = $statuses[0]
            $clean_status = $statuses[1]
            $maidCodeStrings += "{`"status_code`":`"$input_code`",`"occupied_status`":`"$occupied_status`",`"clean_status`":`"$clean_status`"}"
        }

        $reportResult += "Found the following HSKP Codes from $computer`:`n"

        $reportResult += $filteredMaidCodes | Format-Table | Out-String

        $maidCodes = $maidCodeStrings -join ","
    }
    else {
        $maidCodes = "{`"status_code`":`"1`",`"occupied_status`":`"OCCUPIED`",`"clean_status`":`"CLEAN`"},{`"status_code`":`"2`",`"occupied_status`":`"OCCUPIED`",`"clean_status`":`"DIRTY`"},{`"status_code`":`"3`",`"occupied_status`":`"VACANT`",`"clean_status`":`"CLEAN`"},{`"status_code`":`"4`",`"occupied_status`":`"VACANT`",`"clean_status`":`"DIRTY`"},{`"status_code`":`"5`",`"occupied_status`":`"OCCUPIED`",`"clean_status`":`"READY`"},{`"status_code`":`"7`",`"occupied_status`":`"VACANT`",`"clean_status`":`"READY`"}" 
        $reportResult += "Unable to connect to or find HSKP from $computer.`nDefault HSKP Codes were set.`n"
    }

    ###########################################################################

    $propertyID = $currentProperty.id

    $session = New-Object Microsoft.PowerShell.Commands.WebRequestSession
    $session.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/133.0.0.0 Safari/537.36 Edg/133.0.0.0"
    $pbxConfigResponse = Invoke-WebRequest -UseBasicParsing -Uri "$($currentProperty.region.url)hk-property-interfaces/hotelbrand/properties/$($currentProperty.id)/system-config/pbx" -WebSession $session -Headers $headers
    $pbxConfig = $pbxConfigResponse.Content | ConvertFrom-Json

    if ($($pbxConfig.stream_id).ToString().Length -ne 36) {
        Write-Host "No reservation stream has been assigned to PBX, and therefore this request will fail.`n Please set a stream ID for PBX in HK admin."
        $reportResult += "No reservation stream has been assigned to PBX, and therefore this request will fail.`n Please set a stream ID for PBX in HK admin."
        $progress = $false
    }
    else {
        $progress = $true
    }

    if ($null -eq $($pbxConfig.room_alias_type)) {
        Write-Host "No room alias type has been assigned to PBX, and therefore this request will fail.`n Please set a room alias for PBX in HK admin."
        $reportResult += "No room alias type has been assigned to PBX, and therefore this request will fail.`n Please set a room alias for PBX in HK admin."
        $progress = $false
    }
    else {
        $progress = $true
    }


    if ( $progress) {
        $createHSKPBody = "{`"stream_id`":`"$($pbxConfig.stream_id)`",`"property_id`":`"$propertyID`",`"enabled`":$($pbxConfig.enabled),`"room_alias_type`":`"$($pbxConfig.room_alias_type)`",`"aws_thing_name`":`"pep-prod-$($inncode)-1`",`"room_alias_mappings`":[],`"housekeeping_status_mappings`":[$maidCodes]}"
        $response = Invoke-WebRequest -UseBasicParsing -Uri "$($currentProperty.region.url)hk-property-interfaces/hotelbrand/properties/$($currentProperty.id)/system-config/pbx" -Method "PUT" -WebSession $session -Headers $headers -ContentType "application/json;charset=UTF-8" -Body $createHSKPBody
        $reportResult += "Upload to HK result: " + $response.StatusDescription
        Write-Host "Upload to HK result: "  $response.StatusDescription
    }
    else {
 
    }
    
}

$dateTime = Get-Date -Format "yyyy-MM-dd hh:mm:ss TT"
$reportResult += "`n" + "--------------" + "`n" + "Script execution completed at $dateTime."
$dateTimeFileSafe = Get-Date -Format "yyyy-MM-dd_hh-mm-ss"
$savePath = $savedirectory + "\HSKP Status from OnQ - $dateTimeFileSafe.txt"
$reportResult | Out-File -FilePath $savePath


Write-Output "Script execution completed. Summary saved to $savedirectory"

Read-Host "Press Enter to close the window"
exit
