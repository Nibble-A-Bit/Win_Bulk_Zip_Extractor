# Win_Mass_Extractor - PowerShell Script to Simplify Extracting Multiple ZIP Archives

# Define the function to extract ZIP files from a source folder recursively
function Extract-ZIPFiles {
    param (
        [string]$sourceFolder,
        [string]$outputFolder
    )

    # Verify that the source folder exists before proceeding
    if (-not (Test-Path -Path $sourceFolder -PathType Container)) {
        Write-Warning "Source folder '$sourceFolder' not found."
        return
    }

    # Verify that the output folder exists or create it if it doesn't
    if (-not (Test-Path -Path $outputFolder -PathType Container)) {
        try {
            New-Item -ItemType Directory -Path $outputFolder -ErrorAction Stop | Out-Null
        } catch {
            Write-Error "Failed to create output folder '$outputFolder'. Check your permissions and try again."
            return
        }
    }

    # Get all ZIP files in the source folder and its subfolders
    $zipFiles = Get-ChildItem -Path $sourceFolder -Filter "*.zip" -Recurse

    # Initialize the counter and calculate the total number of archives
    $totalArchives = $zipFiles.Count
    $progressCounter = 0

    # Determine if "Overwrite All" option is selected
    $overwriteAll = $false

    # Iterate through each ZIP file and extract its contents to the output folder
    foreach ($zipFile in $zipFiles) {
        try {
            # Update the counter and calculate the extraction progress as a percentage
            $progressCounter++
            $progressPercentage = ($progressCounter / $totalArchives) * 100

            # Format the progress bar
            $progressBar = '[{0}{1}]' -f ('=' * [Math]::Floor($progressPercentage / 10)), (' ' * (10 - [Math]::Floor($progressPercentage / 10)))

            # Extract the ZIP file
            $extractPath = Join-Path -Path $outputFolder -ChildPath $zipFile.BaseName
            if (-not (Test-Path -Path $extractPath)) {
                [System.IO.Compression.ZipFile]::ExtractToDirectory($zipFile.FullName, $extractPath)
                Write-Host ("Extracting: {0,-60} {1,5:F2}% {2}" -f $zipFile.Name, $progressPercentage, $progressBar)
            } else {
                if (-not $overwriteAll) {
                    do {
                        $action = Read-Host "Duplicate found: '$($zipFile.Name)' already exists in the output folder. Choose an action (O)verwrite, Overwrite (A)ll, (S)kip, (R)ename (by appending copy #):"
                    } while ($action -notin 'O', 'A', 'S', 'R')

                    if ($action -eq 'A') {
                        $overwriteAll = $true
                    }
                }

                switch ($action.ToUpper()) {
                    'O' {
                        Get-ChildItem $extractPath | Remove-Item -Force
                        [System.IO.Compression.ZipFile]::ExtractToDirectory($zipFile.FullName, $extractPath)
                        Write-Host ("Extracting: {0,-60} {1,5:F2}% {2}" -f $zipFile.Name, $progressPercentage, $progressBar)
                    }
                    'A' {
                        Get-ChildItem $extractPath | Remove-Item -Force
                        [System.IO.Compression.ZipFile]::ExtractToDirectory($zipFile.FullName, $extractPath)
                        Write-Host ("Extracting: {0,-60} {1,5:F2}% {2}" -f $zipFile.Name, $progressPercentage, $progressBar)
                    }
                    'S' {
                        Write-Host "Skipped $($zipFile.Name)."
                    }
                    'R' {
                        $copyNumber = 1
                        while (Test-Path -Path $extractPath) {
                            $copyNumber++
                            $extractPath = Join-Path -Path $outputFolder -ChildPath "$($zipFile.BaseName) (Copy $copyNumber)"
                        }
                        [System.IO.Compression.ZipFile]::ExtractToDirectory($zipFile.FullName, $extractPath)
                        Write-Host ("Extracting: {0,-60} {1,5:F2}% {2}" -f $zipFile.Name, $progressPercentage, $progressBar)
                    }
                }
            }
        } catch {
            Write-Error "Failed to extract $($zipFile.Name). Error: $_"
        }
    }
}

# Prompt the user for the source folder and set default as .\
$defaultSourceFolder = ".\"
$sourceFolder = Read-Host "Enter the source folder path (default: $defaultSourceFolder)"
if ([string]::IsNullOrEmpty($sourceFolder)) {
    $sourceFolder = $defaultSourceFolder
}

# Prompt the user for the destination folder and set default as .\extracted
$defaultOutputFolder = ".\extracted"
$outputFolder = Read-Host "Enter the destination folder path (default: $defaultOutputFolder)"
if ([string]::IsNullOrEmpty($outputFolder)) {
    $outputFolder = $defaultOutputFolder
}

# Call the function to extract ZIP files from the specified source folder to the destination folder
Extract-ZIPFiles -sourceFolder $sourceFolder -outputFolder $outputFolder
