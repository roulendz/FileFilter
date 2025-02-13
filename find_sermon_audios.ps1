#region Check for ImportExcel module and import it if available
# Check if the ImportExcel module is installed
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Host "The ImportExcel module is not installed." # Inform the user that the module is missing
    Write-Host "Please install it with the following command:" # Instruct the user how to install the module
    Write-Host "    Install-Module ImportExcel -Scope CurrentUser" # Provide the installation command
    return # Exit the script if the module is not installed
} else {
    Import-Module ImportExcel # Import the module if it is installed
}
#endregion

#region Define Show-Menu function for interactive menu selection
function Show-Menu {
    param (
        [string]$Title,           # Title to display on the menu
        [string[]]$Options        # Array of options to display in the menu
    )
    $selectedIndex = 0           # Initialize the selected index to 0
    $key = $null                 # Initialize the key variable

    do {
        Clear-Host              # Clear the host screen for a fresh menu display
        Write-Host $Title       # Display the menu title

        for ($i = 0; $i -lt $Options.Length; $i++) {
            if ($i -eq $selectedIndex) {
                Write-Host "> $($Options[$i])" -ForegroundColor Cyan # Highlight the currently selected option
            } else {
                Write-Host "  $($Options[$i])" # Display unselected options
            }
        }

        $key = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") # Read user key input without echoing it

        switch ($key.VirtualKeyCode) {
            38 { if ($selectedIndex -gt 0) { $selectedIndex-- } }  # Up arrow key: decrease index if not at the top
            40 { if ($selectedIndex -lt $Options.Length - 1) { $selectedIndex++ } }  # Down arrow key: increase index if not at the bottom
            13 { break }  # Enter key: break the loop to select the option
        }
    } while ($key.VirtualKeyCode -ne 13) # Continue loop until Enter is pressed

    return $Options[$selectedIndex] # Return the selected option
}
#endregion

#region Define Get-VariantRegex function to create regex pattern with diacritical variants
function Get-VariantRegex {
    param (
        [string]$inputString  # The input search string from the user
    )
    $pattern = ""
    foreach ($char in $inputString.ToCharArray()) {
        switch ($char) {
            'a' { $pattern += "[aā]" }
            'A' { $pattern += "[AĀ]" }
            'c' { $pattern += "[cč]" }
            'C' { $pattern += "[CČ]" }
            'e' { $pattern += "[eē]" }
            'E' { $pattern += "[EĒ]" }
            'g' { $pattern += "[gģ]" }
            'G' { $pattern += "[GĢ]" }
            'i' { $pattern += "[iī]" }
            'I' { $pattern += "[IĪ]" }
            'k' { $pattern += "[kķ]" }
            'K' { $pattern += "[KĶ]" }
            'l' { $pattern += "[lļ]" }
            'L' { $pattern += "[LĻ]" }
            'n' { $pattern += "[nņ]" }
            'N' { $pattern += "[NŅ]" }
            's' { $pattern += "[sš]" }
            'S' { $pattern += "[SŠ]" }
            'u' { $pattern += "[uū]" }
            'U' { $pattern += "[UŪ]" }
            'z' { $pattern += "[zž]" }
            'Z' { $pattern += "[ZŽ]" }
            default { $pattern += [regex]::Escape($char) }
        }
    }
    return "(?i)$pattern"  # Add case-insensitive flag
}
#endregion


#region Select drive to scan for files
$drives = Get-PSDrive -PSProvider FileSystem | Select-Object -ExpandProperty Root
$Drive = Show-Menu -Title 'Select a Drive to scan for .wav and .mp3 files' -Options $drives
#endregion

#region Select search option (by day or by filename) and get corresponding search criteria
$SearchOptions = @("Search by Day", "Search by File Name")
$SearchOption = Show-Menu -Title 'Select Search Option' -Options $SearchOptions

[string]$DayOfWeek = $null
[string]$SearchText = $null

if ($SearchOption -eq "Search by Day") {
    $DaysOfWeek = [System.Enum]::GetNames([System.DayOfWeek])
    $DayOfWeek = Show-Menu -Title 'Select a Day of The Week' -Options $DaysOfWeek
    if (-not $DayOfWeek) {
        Write-Error "No day selected. Exiting script."
        exit
    }
} elseif ($SearchOption -eq "Search by File Name") {
    $SearchText = Read-Host "Enter text to search in file name or path"
    if (-not $SearchText) {
        Write-Error "No search text provided. Exiting script."
        exit
    }
}
#endregion

#region Prepare output folder and file
$ScriptFolder = Split-Path -Parent $MyInvocation.MyCommand.Definition
$OutputFolder = Join-Path -Path $ScriptFolder -ChildPath "Output"
if (-not (Test-Path -Path $OutputFolder)) {
    New-Item -ItemType Directory -Path $OutputFolder | Out-Null
}

# Determine search descriptor based on search type
[string]$SearchDescriptor = if ($SearchOption -eq "Search by Day") {
    $DayOfWeek
} else {
    $SearchText
}

# Remove invalid file name characters
$SafeDescriptor = [regex]::Replace($SearchDescriptor, '[\\\/\:\*\?\"\<\>\|]', '_')

# Prepare file name with descriptor
$DriveName = $Drive.Substring(0, $Drive.Length - 2)
$timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$OutputFile = Join-Path -Path $OutputFolder -ChildPath "$DriveName ($SafeDescriptor) $timestamp.txt"
#endregion

#region Initialize the output file with header line
"Day,FileName,DateCreated,DateModified,DateLastAccessed,Path" | Out-File -FilePath $OutputFile
#endregion

#region Set culture for date formatting
[System.Threading.Thread]::CurrentThread.CurrentCulture = [System.Globalization.CultureInfo]::GetCultureInfo("lv-LV")
#endregion

#region Collect file details
$fileDetails = Get-ChildItem -Path "$Drive" -Include *.wav,*.mp3 -Recurse -ErrorAction SilentlyContinue | ForEach-Object {
    $dayFromName = if ($_.Name -match "^(\d{4}-\d{2}-\d{2})\s") {
        [datetime]::ParseExact($matches[1], 'yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture).ToString("dddd", [System.Globalization.CultureInfo]::GetCultureInfo("lv-LV"))
    } else {
        "Unknown"
    }
    [PSCustomObject]@{
        Day = $dayFromName
        Name = $_.Name
        CreationTime = $_.CreationTime
        LastWriteTime = $_.LastWriteTime
        LastAccessTime = $_.LastAccessTime
        FullName = $_.FullName
        FormattedCreationTime = $_.CreationTime.ToString("yyyy-MM-ddTHH:mm:ss (dddd)")
        FormattedLastWriteTime = $_.LastWriteTime.ToString("yyyy-MM-ddTHH:mm:ss (dddd)")
        FormattedLastAccessTime = $_.LastAccessTime.ToString("yyyy-MM-ddTHH:mm:ss (dddd)")
    }
}
#endregion

#region Define Remove-Diacritics function to simplify strings
function Remove-Diacritics {
    param (
        [string]$inputString
    )
    # Normalize the string to Form D (NFD), which decomposes characters like ē to e + ˉ
    $normalized = $inputString.Normalize([System.Text.NormalizationForm]::FormD)
    # Remove all non-spacing marks (diacritics)
    $cleaned = -join ($normalized.ToCharArray() | Where-Object { [System.Globalization.CharUnicodeInfo]::GetUnicodeCategory($_) -ne 'NonSpacingMark' })
    return $cleaned
}
#endregion

#region Filter file details
if ($SearchOption -eq "Search by Day") {
    $filteredFileDetails = $fileDetails | Where-Object {
        $_.CreationTime.DayOfWeek.ToString() -eq $DayOfWeek -or $_.Day -eq $DayOfWeek
    }
} elseif ($SearchOption -eq "Search by File Name") {
    $searchPattern = Remove-Diacritics -inputString $SearchText
    $filteredFileDetails = $fileDetails | Where-Object {
        # Normalize the name and path for comparison
        $cleanName = Remove-Diacritics -inputString $_.Name
        $cleanPath = Remove-Diacritics -inputString $_.FullName
        # Use wildcard matching
        $cleanName -like "*$searchPattern*" -or $cleanPath -like "*$searchPattern*"
    }
}
#endregion


#region Sort and write results
$sortedFileDetails = $filteredFileDetails | Sort-Object -Property CreationTime

foreach ($file in $sortedFileDetails) {
    $line = "$($file.Day),$($file.Name),$($file.FormattedCreationTime),$($file.FormattedLastWriteTime),$($file.FormattedLastAccessTime),$($file.FullName)"
    $line | Out-File -FilePath $OutputFile -Append
}
Write-Output "Scan complete. Check '$OutputFile' for the results."
#endregion

#region Convert to Excel
$ExcelFile = [System.IO.Path]::ChangeExtension($OutputFile, ".xlsx")
Import-Csv $OutputFile | Export-Excel -Path $ExcelFile
Write-Output "Excel file created. Check '$ExcelFile' for the results."
#endregion
