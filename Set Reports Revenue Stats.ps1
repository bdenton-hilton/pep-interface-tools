param (
    [string]$inncodes,
    [string]$ssofile,
    [string]$savedirectory
)


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

write-host "$inncodes"

$inncodeArray = @()
$inncodeArray = $inncodes -split "," | ForEach-Object { $_.Trim() }

foreach ($inncode in $inncodeArray){
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

$currentProperty = $sso_login.properties | Where-Object { $_.code -eq $inncode }

Write-Host "`nCurrent Property: $inncode"

$propertyID = $currentProperty.id

$session = New-Object Microsoft.PowerShell.Commands.WebRequestSession
$session.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/133.0.0.0 Safari/537.36 Edg/133.0.0.0"


$chargeTypesReseponse = Invoke-WebRequest -UseBasicParsing -Uri "$($currentProperty.region.url)v4/hotelbrand/properties/$($currentProperty.id)/charge-types" -WebSession $session -Headers $headers
$chargeTypes = $chargeTypesReseponse.Content | ConvertFrom-Json

write-host "Please enter go live date in YYYY-MM-DD"
$validDate = Get-ValidDate
$validDate = $validDate.Trim()
$createRevStatsBody = "{`"charge_type_ids`":[`"$($($chargeTypes | Where-Object { $_.code -eq "ENSRR" }).id)`",`"$($($chargeTypes | Where-Object { $_.code -eq "RR" }).id)`",`"$($($chargeTypes | Where-Object { $_.code -eq "RRA" }).id)`",`"$($($chargeTypes | Where-Object { $_.code -eq "NSRR" }).id)`",`"$($($chargeTypes | Where-Object { $_.code -eq "PEC" }).id)`",`"$($($chargeTypes | Where-Object { $_.code -eq "HHRR" }).id)`",`"$($($chargeTypes | Where-Object { $_.code -eq "MRR" }).id)`",`"$($($chargeTypes | Where-Object { $_.code -eq "RTC" }).id)`",`"$($($chargeTypes | Where-Object { $_.code -eq "ADRR" }).id)`"],`"other_room_revenue_charge_type_ids`":[`"$($($chargeTypes | Where-Object { $_.code -eq "NSCF" }).id)`"],`"payment_type_ids`":[],`"calculation_type`":`"REVENUE`",`"exclude_adjustments`":false,`"id`":null,`"created_at`":null,`"updated_at`":null,`"deleted_at`":null}"

Invoke-WebRequest -UseBasicParsing -Uri "$($currentProperty.region.url)v4/hotelbrand/properties/$($currentProperty.id)/calculation-config?calculation_type=REVENUE&start_date=$($validDate)%22" -Method "PUT" -WebSession $session -Headers $headers -ContentType "application/json" -Body $createRevStatsBody
}

# Add your logic to handle the SSO file and other operations here
Write-Output "Script execution completed."

Read-Host "Press Enter to close the window"
exit
