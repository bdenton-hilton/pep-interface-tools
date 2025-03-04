param (
    [string]$inncodes,
    [string]$ssofile,
    [string]$savedirectory
)

function Replace-ReportID {
    param (
        [string]$inputString
    )
    
    # Find all instances of -replaceme-
    $pattern = "-replaceme-"
    $matches = [regex]::Matches($inputString, $pattern)
    
    # Replace each instance with a unique UUID
    foreach ($match in $matches) {
        $uuid = [guid]::NewGuid().ToString()
        $inputString = $inputString -replace [regex]::Escape($match.Value), $uuid
    }

    return $inputString
}

function Get-ValidDate {
    while ($true) {
        # Prompt the user to enter a date
        $inputDate = Read-Host "Please enter a date in YYYY-MM-DD format"

        # Replace different delimiters with hyphens
        $inputDate = $inputDate -replace '[./]', '-'

        # Split the input date into components
        $dateParts = $inputDate -split '-'

        # Check if the input has three parts (year, month, day)
        if ($dateParts.Length -ne 3) {
            Write-Host "Invalid format. Please enter the date in YYYY-MM-DD format."
            continue
        }

        # Extract year, month, and day
        $year = $dateParts[0]
        $month = $dateParts[1]
        $day = $dateParts[2]

        # Correct two-digit year
        if ($year.Length -eq 2) {
            $year = "20$year"
        }

        # Convert to integers
        $year = [int]$year
        $month = [int]$month
        $day = [int]$day

        # Validate month and day
        if ($month -lt 1 -or $month -gt 12) {
            Write-Host "Invalid month. Please enter a valid month (01-12)."
            continue
        }

        if ($day -lt 1 -or $day -gt 31) {
            Write-Host "Invalid day. Please enter a valid day (01-31)."
            continue
        }

        # Create a DateTime object
        try {
            $date = [datetime]::new($year, $month, $day)
        } catch {
            Write-Host "Invalid date. Please enter a valid date."
            continue
        }

        # Check if the date is in the future
        if ($date -le (Get-Date)) {
            Write-Host "The date must be in the future. Please enter a future date."
            continue
        }

        # Return the valid date as a string in YYYY-MM-DD format
        return $date.ToString("yyyy-MM-dd")
    }
}

#UN-DELIMIT INNCODES

write-host "$inncodes"



$inncodeArray = @()
$inncodeArray = $inncodes -split "," | ForEach-Object { $_.Trim() }


#IMPORT SSO LOGON


$sso_login = ConvertFrom-Json -InputObject (Get-Content -Path $ssofile -Raw)

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

$validDate = Get-ValidDate

foreach ($inncode in $inncodeArray){

$currentProperty = $sso_login.properties | Where-Object { $_.code -eq $inncode }

Write-Host "`nCurrent Property: $inncode"


$reportPacketUUID = [guid]::NewGuid().ToString()
$propertyID = $currentProperty.id
$createPacketBody = "{`"name`":`"Night Audit Report Queue`",`"code`":`"NA`",`"export_type`":`"FILE`",`"active`":false,`"type`":`"COMBINED`",`"subject`":`"$inncode Night Audit Reports`",`"id`":`"$reportPacketUUID`",`"auto_printing`":false,`"sftp_config_id`":null,`"property_id`":`"$propertyID`",`"items`":[{`"report_id`":`"market-segment-summary`",`"template_name`":`"Market Segment Summary`",`"filters`":[{`"id`":`"DATE_FROM`",`"type`":`"date`",`"selected_value`":`"today-1`"},{`"id`":`"DATE_TO`",`"type`":`"date`",`"selected_value`":`"today-1`"}],`"collection_id`":`"$reportPacketUUID`",`"export_type`":{`"excel`":false,`"pdf`":true,`"csv`":null,`"file_name`":null},`"layout`":`"LANDSCAPE`",`"type`":`"DYNAMIC`",`"subject`":null,`"preferred_columns`":null,`"report_file_name`":null,`"id`":`"-replaceme-`",`"created_at`":null,`"updated_at`":null,`"deleted_at`":null},{`"report_id`":`"direct-bill-aging`",`"template_name`":`"Direct Bill Aging`",`"filters`":[{`"id`":`"DATE`",`"type`":`"date`",`"selected_value`":`"today-1`"}],`"collection_id`":`"$reportPacketUUID`",`"export_type`":{`"excel`":false,`"pdf`":true,`"csv`":null,`"file_name`":null},`"layout`":`"LANDSCAPE`",`"type`":`"DYNAMIC`",`"subject`":null,`"preferred_columns`":null,`"report_file_name`":null,`"id`":`"-replaceme-`",`"created_at`":null,`"updated_at`":null,`"deleted_at`":null},{`"report_id`":`"revenue-activity`",`"template_name`":`"Revenue Activity`",`"filters`":[{`"id`":`"DATE_FROM`",`"type`":`"date`",`"selected_value`":`"today-1`"},{`"id`":`"DATE_TO`",`"type`":`"date`",`"selected_value`":`"today-1`"}],`"collection_id`":`"$reportPacketUUID`",`"export_type`":{`"excel`":false,`"pdf`":true,`"csv`":null,`"file_name`":null},`"layout`":`"LANDSCAPE`",`"type`":`"DYNAMIC`",`"subject`":null,`"preferred_columns`":null,`"report_file_name`":null,`"id`":`"-replaceme-`",`"created_at`":null,`"updated_at`":null,`"deleted_at`":null},{`"report_id`":`"HotelStatistics`",`"template_name`":`"Hotel Statistics`",`"filters`":[{`"id`":`"DATE`",`"type`":`"date`",`"selected_value`":`"today-1`"}],`"collection_id`":`"$reportPacketUUID`",`"export_type`":{`"excel`":false,`"pdf`":true,`"csv`":null,`"file_name`":null},`"layout`":`"LANDSCAPE`",`"type`":`"STATIC`",`"subject`":null,`"preferred_columns`":null,`"report_file_name`":null,`"id`":`"-replaceme-`",`"created_at`":null,`"updated_at`":null,`"deleted_at`":null},{`"report_id`":`"in-house-guest-folio-balances`",`"template_name`":`"In House Guest Folio Balances`",`"filters`":[{`"id`":`"DATE`",`"type`":`"date`",`"selected_value`":`"today-1`"}],`"collection_id`":`"$reportPacketUUID`",`"export_type`":{`"excel`":false,`"pdf`":true,`"csv`":null,`"file_name`":null},`"layout`":`"LANDSCAPE`",`"type`":`"DYNAMIC`",`"subject`":null,`"preferred_columns`":null,`"report_file_name`":null,`"id`":`"-replaceme-`",`"created_at`":null,`"updated_at`":null,`"deleted_at`":null},{`"report_id`":`"payment-activity`",`"template_name`":`"Payment Activity`",`"filters`":[{`"id`":`"DATE_FROM`",`"type`":`"date`",`"selected_value`":`"today-1`"},{`"id`":`"DATE_TO`",`"type`":`"date`",`"selected_value`":`"today-1`"}],`"collection_id`":`"$reportPacketUUID`",`"export_type`":{`"excel`":false,`"pdf`":true,`"csv`":null,`"file_name`":null},`"layout`":`"LANDSCAPE`",`"type`":`"DYNAMIC`",`"subject`":null,`"preferred_columns`":null,`"report_file_name`":null,`"id`":`"-replaceme-`",`"created_at`":null,`"updated_at`":null,`"deleted_at`":null},{`"report_id`":`"adjustment-activity`",`"template_name`":`"Adjustments and Refunds Activity`",`"filters`":[{`"id`":`"DATE_FROM`",`"type`":`"date`",`"selected_value`":`"today-1`"},{`"id`":`"DATE_TO`",`"type`":`"date`",`"selected_value`":`"today-1`"}],`"collection_id`":`"$reportPacketUUID`",`"export_type`":{`"excel`":false,`"pdf`":true,`"csv`":null,`"file_name`":null},`"layout`":`"LANDSCAPE`",`"type`":`"DYNAMIC`",`"subject`":null,`"preferred_columns`":null,`"report_file_name`":null,`"id`":`"-replaceme-`",`"created_at`":null,`"updated_at`":null,`"deleted_at`":null},{`"report_id`":`"closed-folios-balance`",`"template_name`":`"Closed Folio Balances`",`"filters`":[{`"id`":`"DATE`",`"type`":`"date`",`"selected_value`":`"today-1`"}],`"collection_id`":`"$reportPacketUUID`",`"export_type`":{`"excel`":false,`"pdf`":true,`"csv`":null,`"file_name`":null},`"layout`":`"LANDSCAPE`",`"type`":`"DYNAMIC`",`"subject`":null,`"preferred_columns`":null,`"report_file_name`":null,`"id`":`"-replaceme-`",`"created_at`":null,`"updated_at`":null,`"deleted_at`":null},{`"report_id`":`"direct-bill-ledger`",`"template_name`":`"Direct Bill Ledger Details`",`"filters`":[{`"id`":`"DATE`",`"type`":`"date`",`"selected_value`":`"today-1`"}],`"collection_id`":`"$reportPacketUUID`",`"export_type`":{`"excel`":false,`"pdf`":true,`"csv`":null,`"file_name`":null},`"layout`":`"LANDSCAPE`",`"type`":`"DYNAMIC`",`"subject`":null,`"preferred_columns`":null,`"report_file_name`":null,`"id`":`"-replaceme-`",`"created_at`":null,`"updated_at`":null,`"deleted_at`":null},{`"report_id`":`"advance-deposit-activity`",`"template_name`":`"Advance Deposit Activity`",`"filters`":[{`"id`":`"DATE`",`"type`":`"date`",`"selected_value`":`"today-1`"}],`"collection_id`":`"$reportPacketUUID`",`"export_type`":{`"excel`":false,`"pdf`":true,`"csv`":null,`"file_name`":null},`"layout`":`"LANDSCAPE`",`"type`":`"DYNAMIC`",`"subject`":null,`"preferred_columns`":null,`"report_file_name`":null,`"id`":`"-replaceme-`",`"created_at`":null,`"updated_at`":null,`"deleted_at`":null},{`"report_id`":`"AllCharges`",`"template_name`":`"All Charges Report`",`"filters`":[{`"id`":`"DATE_FROM`",`"type`":`"date`",`"selected_value`":`"today-1`"},{`"id`":`"DATE_TO`",`"type`":`"date`",`"selected_value`":`"today-1`"},{`"id`":`"SHOW_REMARKS`",`"type`":`"string`",`"selected_value`":`"YES`"},{`"id`":`"ALL_CHARGE_TYPE_ID`",`"type`":`"string`",`"selected_value`":`"all`"}],`"collection_id`":`"$reportPacketUUID`",`"export_type`":{`"excel`":false,`"pdf`":true,`"csv`":null,`"file_name`":null},`"layout`":`"LANDSCAPE`",`"type`":`"STATIC`",`"subject`":null,`"preferred_columns`":null,`"report_file_name`":null,`"id`":`"-replaceme-`",`"created_at`":null,`"updated_at`":null,`"deleted_at`":null},{`"report_id`":`"house-accounts`",`"template_name`":`"House Accounts`",`"filters`":[{`"id`":`"DATE_FROM`",`"type`":`"date`",`"selected_value`":`"today-1`"},{`"id`":`"DATE_TO`",`"type`":`"date`",`"selected_value`":`"today-1`"}],`"collection_id`":`"$reportPacketUUID`",`"export_type`":{`"excel`":false,`"pdf`":true,`"csv`":null,`"file_name`":null},`"layout`":`"LANDSCAPE`",`"type`":`"DYNAMIC`",`"subject`":null,`"preferred_columns`":null,`"report_file_name`":null,`"id`":`"-replaceme-`",`"created_at`":null,`"updated_at`":null,`"deleted_at`":null},{`"report_id`":`"tax-report`",`"template_name`":`"Tax Report`",`"filters`":[{`"id`":`"DATE_FROM`",`"type`":`"date`",`"selected_value`":`"today-1`"},{`"id`":`"DATE_TO`",`"type`":`"date`",`"selected_value`":`"today-1`"}],`"collection_id`":`"$reportPacketUUID`",`"export_type`":{`"excel`":false,`"pdf`":true,`"csv`":null,`"file_name`":null},`"layout`":`"LANDSCAPE`",`"type`":`"DYNAMIC`",`"subject`":null,`"preferred_columns`":null,`"report_file_name`":null,`"id`":`"-replaceme-`",`"created_at`":null,`"updated_at`":null,`"deleted_at`":null},{`"report_id`":`"rate-report`",`"template_name`":`"Rate Report`",`"filters`":[{`"id`":`"DATE`",`"type`":`"date`",`"selected_value`":`"today-1`"},{`"id`":`"status`",`"type`":`"string`",`"selected_value`":`"CHECKED_IN`"}],`"collection_id`":`"$reportPacketUUID`",`"export_type`":{`"excel`":false,`"pdf`":true,`"csv`":null,`"file_name`":null},`"layout`":`"LANDSCAPE`",`"type`":`"DYNAMIC`",`"subject`":null,`"preferred_columns`":null,`"report_file_name`":null,`"id`":`"-replaceme-`",`"created_at`":null,`"updated_at`":null,`"deleted_at`":null},{`"report_id`":`"house-account-balances`",`"template_name`":`"House Account Folio Balances`",`"filters`":[{`"id`":`"DATE`",`"type`":`"date`",`"selected_value`":`"today-1`"}],`"collection_id`":`"$reportPacketUUID`",`"export_type`":{`"excel`":false,`"pdf`":true,`"csv`":null,`"file_name`":null},`"layout`":`"LANDSCAPE`",`"type`":`"DYNAMIC`",`"subject`":null,`"preferred_columns`":null,`"report_file_name`":null,`"id`":`"-replaceme-`",`"created_at`":null,`"updated_at`":null,`"deleted_at`":null},{`"report_id`":`"maintenance-activity`",`"template_name`":`"Maintenance Activity`",`"filters`":[{`"id`":`"DATE_FROM`",`"type`":`"date`",`"selected_value`":`"today-1`"},{`"id`":`"DATE_TO`",`"type`":`"date`",`"selected_value`":`"today-1`"}],`"collection_id`":`"$reportPacketUUID`",`"export_type`":{`"excel`":false,`"pdf`":true,`"csv`":null,`"file_name`":null},`"layout`":`"LANDSCAPE`",`"type`":`"DYNAMIC`",`"subject`":null,`"preferred_columns`":null,`"report_file_name`":null,`"id`":`"-replaceme-`",`"created_at`":null,`"updated_at`":null,`"deleted_at`":null},{`"report_id`":`"clerk-shift`",`"template_name`":`"Clerk Shift`",`"filters`":[{`"id`":`"DATE_FROM`",`"type`":`"date`",`"selected_value`":`"today-1`"},{`"id`":`"DATE_TO`",`"type`":`"date`",`"selected_value`":`"today-1`"},{`"id`":`"user_ids`",`"type`":`"string`",`"selected_value`":`"ALL`"}],`"collection_id`":`"$reportPacketUUID`",`"export_type`":{`"excel`":false,`"pdf`":true,`"csv`":null,`"file_name`":null},`"layout`":`"LANDSCAPE`",`"type`":`"DYNAMIC`",`"subject`":null,`"preferred_columns`":null,`"report_file_name`":null,`"id`":`"-replaceme-`",`"created_at`":null,`"updated_at`":null,`"deleted_at`":null},{`"report_id`":`"FinalAuditV2`",`"template_name`":`"Final Audit`",`"filters`":[{`"id`":`"DATE`",`"type`":`"date`",`"selected_value`":`"today-1`"}],`"collection_id`":`"$reportPacketUUID`",`"export_type`":{`"excel`":false,`"pdf`":true,`"csv`":null,`"file_name`":null},`"layout`":`"LANDSCAPE`",`"type`":`"STATIC`",`"subject`":null,`"preferred_columns`":null,`"report_file_name`":null,`"id`":`"-replaceme-`",`"created_at`":null,`"updated_at`":null,`"deleted_at`":null},{`"report_id`":`"rate-override`",`"template_name`":`"Rate Override`",`"filters`":[{`"id`":`"DATE`",`"type`":`"date`",`"selected_value`":`"today-1`"}],`"collection_id`":`"$reportPacketUUID`",`"export_type`":{`"excel`":false,`"pdf`":true,`"csv`":null,`"file_name`":null},`"layout`":`"LANDSCAPE`",`"type`":`"DYNAMIC`",`"subject`":null,`"preferred_columns`":null,`"report_file_name`":null,`"id`":`"-replaceme-`",`"created_at`":null,`"updated_at`":null,`"deleted_at`":null},{`"report_id`":`"all-transactions`",`"template_name`":`"All Transactions`",`"filters`":[{`"id`":`"DATE_FROM`",`"type`":`"date`",`"selected_value`":`"today-1`"},{`"id`":`"DATE_TO`",`"type`":`"date`",`"selected_value`":`"today-1`"},{`"id`":`"transaction_type`",`"type`":`"string`",`"selected_value`":`"ALL`"},{`"id`":`"user_ids`",`"type`":`"string`",`"selected_value`":`"ALL`"}],`"collection_id`":`"$reportPacketUUID`",`"export_type`":{`"excel`":false,`"pdf`":true,`"csv`":null,`"file_name`":null},`"layout`":`"LANDSCAPE`",`"type`":`"DYNAMIC`",`"subject`":null,`"preferred_columns`":null,`"report_file_name`":null,`"id`":`"-replaceme-`",`"created_at`":null,`"updated_at`":null,`"deleted_at`":null}]}"

$packetBody = Replace-ReportID -inputString $createPacketBody

$session = New-Object Microsoft.PowerShell.Commands.WebRequestSession
$session.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/133.0.0.0 Safari/537.36 Edg/133.0.0.0"

Invoke-WebRequest -UseBasicParsing -Uri "$($currentProperty.region.url)v4-reports/hotelbrand/properties/$($currentProperty.id)/reports/collections" -Method "PUT" -WebSession $session -Headers $headers -ContentType "application/json" -Body "$packetBody"


$gmemail = $inncode + "_gm@hilton.com"
$startdate = ($validDate.Trim() + "T05:00:00Z")

Start-sleep -Milliseconds 1000

Invoke-WebRequest -UseBasicParsing -Uri "$($currentProperty.region.url)v4/hotelbrand/properties/$($currentProperty.id)/email-schedule-reports" -Method "POST" -WebSession $session -Headers $headers -ContentType "application/json" -Body "{`"email`":`"$gmemail`",`"time`":`"$startdate`",`"offset`":1,`"offset_type`":`"DAILY`",`"sftp_config_id`":null,`"property_id`":`"$($currentProperty.id)`",`"report_collection_id`":`"$reportPacketUUID`"}"
}

# Add your logic to handle the SSO file and other operations here
Write-Output "Script execution completed."

Read-Host "Press Enter to close the window"
exit
