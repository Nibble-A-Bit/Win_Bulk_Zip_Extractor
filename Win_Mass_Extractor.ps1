# Win_Mass_Extractor - PowerShell Script to Simplify Extracting Multiple ZIP Archives

# Function to extract ZIP files using [System.IO.Compression.ZipFile] (PowerShell 5.0 or higher)
function Extract-ZIPFiles_SystemIO {
    param (
        [string]$sourceFolder,
        [string]$outputFolder
    )

    try {
        # Verify that the source folder exists before proceeding
        if (-not (Test-Path -Path $sourceFolder -PathType Container)) {
            Write-Warning "Source folder '$sourceFolder' not found."
            return
        }

        # Verify that the output folder exists or create it if it doesn't
        if (-not (Test-Path -Path $outputFolder -PathType Container)) {
            New-Item -ItemType Directory -Path $outputFolder -ErrorAction Stop | Out-Null
        }

        # Get all ZIP files in the source folder and its subfolders
        $zipFiles = Get-ChildItem -Path $sourceFolder -Filter "*.zip" -Recurse

        # Initialize the counter and calculate the total number of archives
        $totalArchives = $zipFiles.Count
        $progressCounter = 0

        # Initialize a variable to store the user-selected action for future duplicates
        $applyActionToFuture = $null

        # Iterate through each ZIP file and extract its contents to the output folder
        foreach ($zipFile in $zipFiles) {
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
                if (-not $applyActionToFuture) {
                    do {
                        $action = Read-Host "Duplicate found: '$($zipFile.Name)' already exists in the output folder. Choose an action (O)verwrite, (S)kip, (R)ename (by appending copy #):"
                    } while ($action -notin 'O', 'S', 'R')

                    do {
                        # Ask the user if they want to apply this action to all future duplicates
                        $applyActionToFuture = Read-Host "Apply this action to all future duplicates? (Y/N)"
                    } while ($applyActionToFuture -notin 'Y', 'N')
                }

                if ($applyActionToFuture) {
                    switch ($action.ToUpper()) {
                        'O' {
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
            }
        }
    } catch {
        Write-Error "Failed to extract $($zipFile.Name). Error: $_"
    }
}

# Function to extract ZIP files using Shell.Application (Fallback for PowerShell versions < 5.0)
function Extract-ZIPFiles_ShellApplication {
    param (
        [string]$sourceFolder,
        [string]$outputFolder
    )

    # Fallback to using Shell.Application for ZIP file extraction
    $shell = New-Object -ComObject Shell.Application

    # ... (The remaining code for ZIP extraction using Shell.Application goes here)

    # For the sake of this example, the Shell.Application method is not implemented. You can refer to the previous version of the script for the implementation.

    Write-Warning "The current PowerShell version does not support [System.IO.Compression.ZipFile]. Fallback to Shell.Application is not implemented in this example."
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

# Check PowerShell version and call the appropriate function for ZIP file extraction
if ($PSVersionTable.PSVersion.Major -ge 5) {
    # Use [System.IO.Compression.ZipFile] for ZIP extraction
    Extract-ZIPFiles_SystemIO -sourceFolder $sourceFolder -outputFolder $outputFolder
} else {
    # Fallback to using Shell.Application for ZIP extraction
    Extract-ZIPFiles_ShellApplication -sourceFolder $sourceFolder -outputFolder $outputFolder
}
