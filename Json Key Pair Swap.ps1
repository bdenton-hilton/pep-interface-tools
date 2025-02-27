# Load the JSON data from a file
$jsonContent = Get-Content -Path "C:\Users\Brady Denton\Desktop\MEMPE\MEMPE_IFC_V5_DoorLockSetup.json" -Raw
$jsonData = ConvertFrom-Json -InputObject $jsonContent

# Check if the LOCKTYPES object exists
if ($jsonData.LOCKTYPES -ne $null) {
    # Create an ordered hashtable to store the swapped key-value pairs
    $swappedLockTypes = [ordered]@{}

    # Iterate through each key-value pair in LOCKTYPES and swap them
    foreach ($key in $jsonData.LOCKTYPES.PSObject.Properties.Name) {
        if ($key -as [int]) {
            $value = $jsonData.LOCKTYPES.$key
            $swappedLockTypes["$value"] = $key  # Convert value to string
        } else {
            $swappedLockTypes[$key] = $jsonData.LOCKTYPES.$key  # Retain original key-value pair
        }
    }

    # Update the LOCKTYPES object in the original JSON data
    $jsonData.LOCKTYPES = $swappedLockTypes
}

# Convert the updated JSON data back to JSON
$updatedJson = $jsonData | ConvertTo-Json -Depth 100

# Save the swapped JSON data to a file
Set-Content -Path "C:\Users\Brady Denton\Desktop\MEMPE\MEMPE_IFC_V5_DoorLockSetup_Swapped.json" -Value $updatedJson
