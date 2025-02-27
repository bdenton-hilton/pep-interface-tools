#Active Json Sorter
$jsonFile = Get-Content -Path "C:\Users\Brady Denton\Desktop\MEMPE\MEMPE_IFC_V5_DoorLockSetup_Swapped.json" -Raw
$jsonContent = ConvertFrom-Json -InputObject $jsonFile

# Access the "locktypes" object using dot notation
$locktypesObject = $jsonContent.locktypes

#creates a list of all key types (more so, make a list of the names of all items under the [LOCKTYPE] header)
$listA = @()
foreach ($name in $locktypesObject.PSObject.Properties) {
    $listA += ($name.name).Trim()
}
#creates a list of settings, permissions and additional object/header names. 
$listB = @()
foreach ($name in $jsonContent.PSObject.Properties) {
    $listB += $name.name
}

# Foreach loop over $listA
foreach ($item in $listA) {
    $temp_list = @()

    $digit_handler = $item
    #handler for Saflok SAFLOCK and VING systems
    if ($item -match '\d') {
        $word = $item -split '\d', 2 | Select-Object -First 1
        $item = $word
    }

    # Foreach loop over $listB to find matches
    foreach ($bItem in $listB) {

        # Convert both items to lowercase for case-insensitive comparison
        $lowerItem = $item.ToLower()
        $lowerBItem = $bItem.ToLower()

        # Check if $lowerItem is a substring of $lowerBItem
        if ($lowerBItem -match $lowerItem) {

            # Check if the word after $lowerItem is "atlas" - handler for Ilco Atlas vs Ilco
            $words = $lowerBItem.Split(' ')
            $indexOfItem = $words.IndexOf($lowerItem)

            if ($indexOfItem -ge 0 -and $indexOfItem -lt ($words.Count - 1)) {
                $nextWord = $words[$indexOfItem + 1]
                if ($nextWord -eq "atlas") {
                    continue
                }
            }

            # Remove $item from $bItem and add the remaining string to the temporary list
            $temp_list += $bItem
        }
    }
    $renamer = $item
    $item = $digit_handler

    if ($temp_list.Count -gt 0) {
        $locktypesObject.$item = @{}
    
        foreach ($object in $temp_list) {
            # Create a nested object within the "locktypes" object
            $renaming_string = ($object -replace $renamer, "" -replace '\s+', ' ').Trim()
            # Check if $bItem is exactly "MESSERSCHMITT"
            if ($object -eq "MESSERSCHMITT") {
                # Update $bItem to "MESSERSCHMITT SETTINGS"
                $renaming_string = "SETTINGS"
            }
            $locktypesObject.$item += @{
                $renaming_string = $jsonContent.$object
            }
        }
    }
}
    
foreach ($object in $listB) {
    # Skip objects with specific names
    if ($object -ne "LOCKTYPES" -and $object -ne "SERVER SETTINGS") {
        $jsonContent.PSObject.Properties.Remove($object)
    }
}
    

# Create a new ordered hashtable to maintain the order of the keys
$orderedJson = [ordered]@{}

# Add SERVER SETTINGS first
$orderedJson['SERVER SETTINGS'] = $jsonContent.'SERVER SETTINGS'

# Then add LOCKTYPES
$orderedJson['LOCKTYPES'] = $jsonContent.LOCKTYPES

# Add other keys if there are any
foreach ($key in $json.PSObject.Properties.Name) {
    if ($key -ne 'SERVER SETTINGS' -and $key -ne 'LOCKTYPES') {
        $orderedJson[$key] = $json[$key]
    }
}

$jsonContent.PSObject.Properties.Remove('Length')

# Convert back to JSON
$orderedJson | ConvertTo-Json -Depth 100 | Set-Content -Path "C:\Users\Brady Denton\Desktop\MEMPE\MEMPE_IFC_V5_DoorLockSetup_Organized.json"
