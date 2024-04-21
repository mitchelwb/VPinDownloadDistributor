# Define the source folder
$sourceFolder = "$env:USERPROFILE\Downloads"

# Define the ROM folder
$romFolder = "C:\vPinball\VisualPinball\VPinMAME\roms"
$altColorFolder = "C:\vPinball\VisualPinball\VPinMAME\altcolor"
$tablesFolder ="C:\vPinball\VisualPinball\Tables"

########## Temporary for Testing ##########
$romFolder = "C:\vPinballTest\VisualPinball\VPinMAME\roms"
$altColorFolder = "C:\vPinballTest\VisualPinball\VPinMAME\altcolor"
$tablesFolder ="C:\vPinballTest\VisualPinball\Tables"
########## Temporary for Testing ##########


## Get all .zip files in the source folder and look for Roms
$zipFiles = Get-ChildItem -Path $sourceFolder -Filter *.zip

# Move Rom files
Clear-Host
$outputString = @"


     ******************************
     *            Roms            *   
     ******************************


"@
Write-Host $outputString

if ($zipFiles.Count -eq 0) {
    Write-Host "No .zip files found in the source folder."
} else {
    $i = 1
    foreach ($zipFile in $zipFiles) {
        Write-Host "  $i. $($zipFile.Name)"
        $i++
    }
    
    do {
        $choiceInput = Read-Host "`nEnter the number of the ROM file you want to move (0 to finish)"
        try{
            $choice = [int]$choiceInput
        }
        catch {
            Write-Host "That's not a number doofus."
            continue
        }

        if ($choice -eq 0) {
            Write-Host "Exiting the script."
            break
        } elseif (($choice -lt 0) -or ($choice -gt $zipFiles.Count)) {
            Write-Host "Invalid choice *$choice*. Please enter a number between 0 and $($zipFiles.Count)."
            Write-Host "Bad Script! $($zipFiles[$choice - 1])"
        } else {
            $selectedZip = $zipFiles[$choice - 1]
            $romPath = Join-Path -Path $romFolder -ChildPath $selectedZip.Name
            Move-Item -Path $selectedZip.FullName -Destination $romPath -Force
            Write-Host "$($selectedZip.Name) moved to $romPath"
        }
    } while ($true)
}

## Now all the remaining .zip files can be unzipped
Clear-Host
$outputString = @"


     ******************************
     *      Tables/Backglass      *   
     ******************************

"@
Write-Host $outputString

# Initialize an empty array to store .zip files
$zipFiles = @()

# Get all .zip files in the source folder
if (Test-Path -Path $sourceFolder) {
    $zipFiles = Get-ChildItem -Path $sourceFolder -Filter *.zip
    #Read-Host "Press Enter to continue..."
    Write-Host "Unzipping files... `n"
}

if ($zipFiles.Count -eq 0) {
    Write-Host "No .zip files found in the source folder to extract."
} else {
    foreach ($zipFile in $zipFiles) {
                
        Write-Host "Unzipping $($zipFile.Name)"

        # Create unzip folder
        $unzipFolder = New-Item -ItemType Directory -Path "$sourceFolder\unzip" -Force

        # Unzip the contents of the zip file to the unzip folder
        Expand-Archive -Path $zipFile.FullName -DestinationPath $unzipFolder.FullName

        # Get files with specified extensions
        $filesToRename = Get-ChildItem -Path $unzipFolder.FullName -Include *.vni, *.crz, *.pal, *.pac -Recurse

        # Rename files
        foreach ($file in $filesToRename) {
            $newName = $zipFile.BaseName + $file.Extension
            Rename-Item -Path $file.FullName -NewName $newName
        }

        # Move files from unzip folder to source folder
        Move-Item -Path "$unzipFolder\*" -Destination $sourceFolder -Force

        # Remove unzip folder
        Remove-Item -Path $unzipFolder -Force -Recurse

    }
}

## Now we can iterate through all the .vpx files and rename them and their corresponding .directb2s files and move them to the Tables

# Get all .vpx files in the Downloads folder
$vpxFiles = Get-ChildItem -Path $sourceFolder -Filter *.vpx

# Loop through each .vpx file
foreach ($vpxFile in $vpxFiles) {
    # Extract the filename without extension
    $vpxFilenameWithoutExtension = $vpxFile.Name -replace '\.vpx$'
    
    # Prompt to update filename
    $newVpxFilename = Read-Host -Prompt ("`nRename $vpxFilenameWithoutExtension.vpx? (Press enter to accept as default)") 
    if ($newVpxFilename -eq ""){
        $newVpxFilename = $vpxFilenameWithoutExtension
    }
    else {
        if ($newVpxFilename -match '(?i)\.vpx$') {
            $newVpxFilename = $newVpxFilename.Substring(0, $newVpxFilename.Length - 4)
        }
        Rename-Item -Path $vpxFile.FullName -NewName "$newVpxFilename.vpx" -Force
    }
    
    # Construct the wildcard pattern to find corresponding .directb2s file
    $directb2sPattern = Join-Path -Path $sourceFolder -ChildPath "*$newVpxFilename*.directb2s"
    
    # Find the corresponding .directb2s file
    $matchingDirectb2sFile = Get-ChildItem -Path $directb2sPattern
    
    # If a matching .directb2s file is found, rename it and move both files to the tables folder
    if ($matchingDirectb2sFile) {
        $newDirectb2sFilename = $newVpxFilename + ".directb2s"
        Rename-Item -Path $matchingDirectb2sFile.FullName -NewName $newDirectb2sFilename -Force
        Write-Output "`nRenamed $($matchingDirectb2sFile.Name) to $newDirectb2sFilename. Moving them to $tablesFolder"
        Move-Item -Path "$sourceFolder\$newDirectb2sFilename" -Destination $tablesFolder -Force
        Move-Item -Path "$sourceFolder\$newVpxFilename.vpx" -Destination $tablesFolder -Force

        # Move-Item -Path (Join-Path -Path $sourceFolder -ChildPath $newDirectb2sFilename.FullName) -Destination $tablesFolder -Force
        # Move-Item -Path (Join-Path -Path $sourceFolder -ChildPath $newVpxFilename.FullName) -Destination $tablesFolder -Force
    } else {
        $moveVpxOnly = Read-Host "No matching .directb2s file found for $($vpxFile.Name). Move vpx file to tables folder? (only 'yes' will work)"
        if($moveVpxOnly -eq "yes"){
            Write-Host "Moving $newVpxFilename.vpx to $tablesFolder"
            Move-Item -Path "$sourceFolder\$newVpxFilename.vpx" -Destination $tablesFolder -Force
        }
    }
}

#Read-Host "`n`nAll vpx files have been moved. Press Enter to continue..."
Clear-Host
$outputString = @"


******************************
*          AltColor          *   
******************************


"@
Write-Host $outputString

#First define the function to update the registry
function Set-RegistryKey {
    param (
        [string]$KeyName
    )

    $keyPath = "HKCU:\Software\Freeware\Visual PinMame\$KeyName"
    Write-Host "Checking $keyPath"

    if (-not (Test-Path $keyPath)) {
        # Create the registry key if it doesn't exist
        Write-Host "Creating $keyPath"
        New-Item -Path $keyPath -Force -WhatIf #| Out-Null
    }

    # Check if the DWORD value exists, if not, create it
    # $keyValueExists = Get-ItemProperty -Path $keyPath -Name "dmd_colorize" -ErrorAction SilentlyContinue
    # if (-not $keyValueExists) {
    #     New-ItemProperty -Path $keyPath -Name "dmd_colorize" -Value 1 -PropertyType DWORD -Force #| Out-Null
    # }
}


## This is where I need to run the script to create folders in the altColor folder and go through my color rom files to choose where they should live

# Initialize an empty array to store the created folders
$createdFolders = @()

# Iterate through zip files in $romFolder and create folders in $altColorFolder
Get-ChildItem -Path $romFolder -Filter *.zip | ForEach-Object {
    $zipName = $_.BaseName
    $folderPath = "$altColorFolder\$zipName"
    if (-not (Test-Path $folderPath)) {
        New-Item -ItemType Directory -Path $folderPath | Out-Null
        #Write-Host "Created folder: $zipName"
        $createdFolders += $folderPath
    }
}

# Enumerate through files in $sourceFolder
#$sourceFiles = Get-ChildItem -Path "$sourceFolder\*" -Recurse -File -Include *.vni, *.crz, *.pal, *.pac
$sourceFiles = Get-ChildItem -Path $sourceFolder -Recurse -File | Where-Object { $_.Extension -match '\.(vni|pal|pac|crz)$' }
Write-Host "`nLooking for color files."
# Iterate through each file and prompt user to select a folder
for ($i = 0; $i -lt $sourceFiles.Count; $i++) {
    $file = $sourceFiles[$i]
    Write-Host "`nChoose destination folder for $($file.FullName):`n"

    # Display numbered list of created folders
    for ($j = 0; $j -lt $createdFolders.Count; $j++) {
        Write-Host " $($j + 1). $($createdFolders[$j])"
    }

    # Prompt user for folder selection
    do {
        $folderChoice = Read-Host "`nEnter folder number (enter '0' to skip)"
        if ($folderChoice -eq "0") {
            Write-Host "Skipped moving $($file.Name)"
            break
        }
        elseif ($folderChoice -eq "M" -or $folderChoice -eq "m") {
            #todo: this needs work
            Write-Host "Additional folders:"
            Get-Folders | ForEach-Object { Write-Host $_ }
        }
    } while ($folderChoice -lt 0 -or $folderChoice -gt $createdFolders.Count -and $folderChoice -ne "M" -and $folderChoice -ne "m")

    # Move file to selected folder
    if ($folderChoice -ne "0" -and $folderChoice -ne "M" -and $folderChoice -ne "m") {
        $selectedFolder = $createdFolders[$folderChoice - 1]
        Move-Item -Path $file.FullName -Destination $selectedFolder -Force
        
        #Set-RegistryKey -KeyName (($selectedFolder -split '\\')[-1])
        $KeyName = ($selectedFolder -split '\\')[-1]
        $keyPath = "HKCU:\Software\Freeware\Visual PinMame\$KeyName"
        Write-Host "Checking $keyPath"
    
        if (-not (Test-Path $keyPath)) {
            # Create the registry key if it doesn't exist
            Write-Host "Creating $keyPath"
            New-Item -Path $keyPath -Force  #| Out-Null
        }
    
        # Check if the DWORD value exists, if not, create it
        $keyValueExists = Get-ItemProperty -Path $keyPath -Name "dmd_colorize" -ErrorAction SilentlyContinue
        if (-not $keyValueExists) {
            Write-Host "Creating dmd_colorize property"
            New-ItemProperty -Path $keyPath -Name "dmd_colorize" -Value 1 -PropertyType DWORD -Force #| Out-Null
        }

        Write-Host "Moved $($file.Name) to $($selectedFolder)"
        Read-Host "Press Enter"
    }

    Clear-Host
    Write-Host $outputString

}

## Now we need to re-name our color rom files to match the conventional standards.
Write-Host "Renaming altColor files appropriately"

# Loop through all folders in the altColor directory
Get-ChildItem -Path $altColorFolder -Directory | ForEach-Object {
    Write-Host "Inspecting folder: $($_.Name)"
    #$filesFound = $false
    #$allFilesNamedCorrectly = $true

    # Loop through all files in the current folder
    Get-ChildItem -Path $_.FullName | ForEach-Object {
        if (-not $_.PSIsContainer) {
            #$filesFound = $true
            $extension = $_.Extension.ToLower()

            # Check if the file extension is .vni, .pal, .pac, or .crz
            if ($extension -eq ".vni" -or $extension -eq ".pal" -or $extension -eq ".pac") {
                # If the file is not named "pin2dmd", set the flag to false
                if ($_.BaseName.ToLower() -ne "pin2dmd") {
                    #$allFilesNamedCorrectly = $false
                    # Copy and rename the file
                    $newFileName = Join-Path -Path $_.Directory.FullName -ChildPath "pin2dmd$extension"
                    if (-not (Test-Path -Path $newFileName)) {
                        # Copy and rename the file
                        Copy-Item -Path $_.FullName -Destination $newFileName
                        Write-Host "Copied and renamed $($_.Name) to $newFileName"
                    }
                    # else {
                    #     "$newFileName already exists"
                    # }
                }
            } elseif ($extension -eq ".crz") {
                # If the file name is different from the parent folder name, set the flag to false
                $parentFolder = Split-Path -Leaf $_.Directory.FullName
                if ($_.BaseName.ToLower() -ne $parentFolder.ToLower()) {
                    #$allFilesNamedCorrectly = $false
                    # Copy and rename the file
                    $newFileName = Join-Path -Path $_.Directory.FullName -ChildPath "$parentFolder$extension"
                    if (-not (Test-Path -Path $newFileName)) {
                        # Copy and rename the file
                        Copy-Item -Path $_.FullName -Destination $newFileName
                        Write-Host "Copied and renamed $($_.Name) to $newFileName"
                    }
                    # else {
                    #     "$newFileName already exists"
                    # }
                }
            }
        }
    }

    # Output appropriate message based on file presence and correctness
    # if (-not $filesFound) {
    #     Write-Host "This folder is empty."
    # } else {
    #     if ($allFilesNamedCorrectly) {
    #         Write-Host "All files in this folder are named correctly."
    #     } else {
    #         Write-Host "This folder contains files with incorrect names."
    #     }
    # }
}

Read-Host "Alt Color file review complete. Press Enter to continue..."

## time to cleanup
Clear-Host
$outputString = @"


     ******************************
     *          Cleanup           *   
     ******************************


"@
Write-Host $outputString


    # Output the files in the source folder
    $files = Get-ChildItem $sourceFolder
    Write-Host "Files currently in the folder:`n`n"
    $files | ForEach-Object { Write-Host $_.Name }

    # Ask if the user wants to clean up the folder
    $response = Read-Host "Do you want to clean up the folder? (only 'yes' will delete files)"

    if ($response.ToLower() -eq "yes") {
        # Permanently delete all files in the source folder
        Remove-Item $sourceFolder\* -Force -Recurse
        Write-Host "All files in the folder have been permanently deleted."
    } else {
        Write-Host "No files have been deleted."
    }

Read-Host "Process complete.  Press enter to close."