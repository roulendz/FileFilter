#region Check for ImportExcel module and import it if available
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Host "The ImportExcel module is not installed."
    Write-Host "Please install it with the following command:"
    Write-Host "    Install-Module ImportExcel -Scope CurrentUser"
    return
} else {
    Import-Module ImportExcel
}
#endregion

#region Updated Show-Menu function with optional default selection
function Show-Menu {
    param (
        [string]$Title,
        [string[]]$Options,
        [bool]$AllowMultiple = $false,
        [int]$DefaultIndex = 0
    )
    $selectedIndexes = @()
    if ($AllowMultiple -and $Options.Length -gt 0) {
        $selectedIndexes = @($DefaultIndex)  # Preselect the first drive
    } elseif (-not $AllowMultiple) {
        $selectedIndexes = @($DefaultIndex) # Preselect single default option
    }

    $currentIndex = if ($Options.Length -gt 0) { $DefaultIndex } else { 0 }
    $key = $null

    do {
        Clear-Host
        Write-Host $Title

        for ($i = 0; $i -lt $Options.Length; $i++) {
            $prefix = if ($selectedIndexes -contains $i) { "[X]" } else { "[ ]" }
            if ($i -eq $currentIndex) {
                Write-Host "> $prefix $($Options[$i])" -ForegroundColor Cyan
            } else {
                Write-Host "  $prefix $($Options[$i])"
            }
        }

        Write-Host "`nUse Up/Down to navigate, Space to select, Enter to confirm."

        $key = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")

        switch ($key.VirtualKeyCode) {
            38 { if ($currentIndex -gt 0) { $currentIndex-- } }
            40 { if ($currentIndex -lt $Options.Length - 1) { $currentIndex++ } }
            32 {
                if ($AllowMultiple) {
                    if ($selectedIndexes -contains $currentIndex) {
                        $selectedIndexes = $selectedIndexes -ne $currentIndex
                    } else {
                        $selectedIndexes += $currentIndex
                    }
                } else {
                    $selectedIndexes = @($currentIndex)
                }
            }
            13 { break }
        }
    } while ($key.VirtualKeyCode -ne 13)

    return $selectedIndexes | ForEach-Object { $Options[$_] }
}
#endregion


#region Select drives to scan for files
$drives = Get-PSDrive -PSProvider FileSystem | Select-Object -ExpandProperty Root
$SelectedDrives = Show-Menu -Title 'Select Drives to scan for .wav and .mp3 files (use Space to select multiple)' -Options $drives -AllowMultiple $true -DefaultIndex 0
if (-not $SelectedDrives) {
    Write-Error "No drives selected. Exiting script."
    exit
}
#endregion

#region Select search option (by day or by filename)
$SearchOptions = @("Search by Day", "Search by File Name")
$defaultIndex = 1  # Set default selection to "Search by File Name"
$SearchOption = Show-Menu -Title 'Select Search Option (only one)' -Options $SearchOptions -AllowMultiple $false -DefaultIndex $defaultIndex
if (-not $SearchOption) {
    $SearchOption = $SearchOptions[$defaultIndex]
}


[string]$DayOfWeek = $null
[string]$SearchText = $null

if ($SearchOption -eq "Search by Day") {
    $DaysOfWeek = [System.Enum]::GetNames([System.DayOfWeek])
    $DayOfWeek = Show-Menu -Title 'Select a Day of The Week' -Options $DaysOfWeek -AllowMultiple $false
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

[string]$SearchDescriptor = if ($SearchOption -eq "Search by Day") { $DayOfWeek } else { $SearchText }
$SafeDescriptor = [regex]::Replace($SearchDescriptor, '[\\/:*?"<>|]', '_')
$DriveNames = ($SelectedDrives -join "_").TrimEnd('\')
$DriveNames = [regex]::Replace($DriveNames, '[\\/:*?"<>|]', '_')
$timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$OutputFile = Join-Path -Path $OutputFolder -ChildPath "$DriveNames ($SafeDescriptor) $timestamp.txt"
#endregion


#region Initialize the output file with header line
"Day,FileName,DateCreated,DateModified,DateLastAccessed,Path" | Out-File -FilePath $OutputFile
#endregion

#region Set culture for date formatting
[System.Threading.Thread]::CurrentThread.CurrentCulture = [System.Globalization.CultureInfo]::GetCultureInfo("lv-LV")
#endregion

#region Collect file details
$fileDetails = @()
foreach ($Drive in $SelectedDrives) {
    $fileDetails += Get-ChildItem -Path "$Drive" -Include *.wav,*.mp3 -Recurse -ErrorAction SilentlyContinue | ForEach-Object {
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
}
#endregion

#region Filter file details
if ($SearchOption -eq "Search by Day") {
    $filteredFileDetails = $fileDetails | Where-Object {
        $_.CreationTime.DayOfWeek.ToString() -eq $DayOfWeek -or $_.Day -eq $DayOfWeek
    }
} elseif ($SearchOption -eq "Search by File Name") {
    $searchPattern = $SearchText
    $filteredFileDetails = $fileDetails | Where-Object {
        $_.Name -like "*$searchPattern*" -or $_.FullName -like "*$searchPattern*"
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
