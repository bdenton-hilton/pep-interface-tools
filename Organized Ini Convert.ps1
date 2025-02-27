function Convert-IniToJson {
    param (
        [string]$IniFilePath
    )

    $orderedJson = [ordered]@{}
    $currentSection = ""

    Get-Content $IniFilePath | ForEach-Object {
        $line = $_.Trim()

        if ($line.StartsWith("[") -and $line.EndsWith("]")) {
            # Section header
            $currentSection = $line.TrimStart('[').TrimEnd(']')
            $orderedJson[$currentSection] = [ordered]@{}
        } elseif ($line -match "^([^=]+)=(.+)$") {
            # Key-value pair
            $key = $Matches[1].Trim()
            $value = $Matches[2].Trim()
            $orderedJson[$currentSection][$key] = $value
        }
    }

    $orderedJson | ConvertTo-Json
}

# Example usage
$jsonOutput = Convert-IniToJson "C:\Users\Brady Denton\Desktop\AMADO\AMADO_IFC_localifs.ini"
$jsonOutput | Out-File "C:\Users\Brady Denton\Desktop\AMADO\AMADO_IFC_localifs.json"