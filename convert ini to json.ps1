# Define the path of the .ini file and the path of the output .json file
$file = "C:\Users\Brady Denton\Desktop\AMADO\AMADO_IFC_localifs.txt"
$jsonFilePath = "C:\Users\Brady Denton\Desktop\AMADO\AMADO_IFC_localifs.json"

function Get-IniContent ($filePath) {
    $ini = @{}
    $section = ""

    switch -regex -file $filePath {
        "^\[(.+)\]" { # Section
            $section = $matches[1]
            $ini[$section] = @{}
        }
        "^(;.*)$" { # Comment (ignore)
        }
        "(.+?)\s*=\s*(.*)" { # Key
            $name, $value = $matches[1..2]
            $ini[$section][$name] = $value
        }
    }

    return $ini
}

# Get the INI content as a hashtable
$iniContent = Get-IniContent $file

# Convert the hashtable to JSON
$json = $iniContent | ConvertTo-Json

# Output the JSON
$json | Out-File -FilePath $jsonFilePath