# Win_Mass_Extractor - PowerShell Script to Simplify Extracting Multiple ZIP Archives

# Function to extract ZIP files using [System.IO.Compression.ZipFile] (PowerShell 5.0 or higher)
function CanIUse-SystemIO {
    try {
        # Check if System.IO.Compression.ZipFile is available by attempting to load the assembly
        [void][Reflection.Assembly]::LoadWithPartialName("System.IO.Compression.FileSystem")
        $zipFileAvailable = $true
    } catch {
        $zipFileAvailable = $false
    }

    return $zipFileAvailable
}
function Extract-ZIPFilesystemIO {
    param (
        [System.IO.FileInfo]$zipFile,
        [string]$extractPath
    )
    if($zipFile -and $extractPath){
        [System.IO.Compression.ZipFile]::ExtractToDirectory($zipFile.FullName, $extractPath)
    } else {throw "Filename or Extraction Path is invalid"}
}

function Extract-ZIPFileShellApp {
    param (
        [System.IO.FileInfo]$zipFile,
        [string]$extractPath,
        [object]$shellApp
    )

     # Verify that $extractPath is not null or empty
     if ([string]::IsNullOrEmpty($extractPath)) {
        Write-Warning "Output folder path is invalid or empty."
        return
    }
    
    # Verify that the output folder exists or create it if it doesn't
    if (-not (Test-Path -Path $extractPath -PathType Container)) {
        New-Item -ItemType Directory -Path $extractPath -ErrorAction Stop | Out-Null
    }
    # Extract the ZIP file using Shell.Application
    $zipFolder = $shellApp.NameSpace($zipFile.FullName)
    $destinationFolder = $shellApp.NameSpace($extractPath)
    $destinationFolder.CopyHere($zipFolder.Items(), 16)  # 16 is the option for 'No Progress Display'

}

function Write-ProgressBar {
    param (
        [string]$actionText,
        [string]$zipFileName,
        [int]$progressCounter,
        [int]$totalArchives
    )
    if ($actionText -and $zipFileName -and $progressCounter -and $totalArchives) {
        $progressPercentage = ($progressCounter / $totalArchives) * 100
    
        # Format the progress bar
        $progressBar = '[{0}{1}]' -f ('=' * [Math]::Floor($progressPercentage / 10)), (' ' * (10 - [Math]::Floor($progressPercentage / 10)))
        Write-Host ("{0,-25}: {1,-60} {2,6:F2}% {3} {4}/{5}" -f $actionText, $zipFile.Name, $progressPercentage, $progressBar, $progressCounter, $totalArchives)
    }
}
function Extract-ZIPFiles {
    param (
        [string]$sourceFolder,
        [string]$outputFolder
    )

    try {
        # Get all ZIP files in the source folder and its subfolders
        $zipFiles = Get-ChildItem -Path $sourceFolder -Filter "*.zip" -Recurse
        $powerShellVersion  = $PSVersionTable.PSVersion.Major
        Write-Host ("Confirmed PowerShell {0}.V" -f $powerShellVersion)

        # Initialize the counter and calculate the total number of archives
        $totalArchives = $zipFiles.Count
        $progressCounter = 0

        # Initialize a variable to store the user-selected action for future duplicates
        $applyActionToFuture = $null
        $systemIOAvail = CanIUse-SystemIO
        if (-not $systemIOAvail) {
            # Create a Shell.Application object to work with ZIP files
            $shellApp = New-Object -ComObject Shell.Application
            # Verify that the Shell.Application object was created successfully
            if ($null -eq $shellApp) {
                Write-Warning "Failed to create Shell.Application object. $_"
                return
    }
        }

        # Iterate through each ZIP file and extract its contents to the output folder
        foreach ($zipFile in $zipFiles) {
            # Update the counter and calculate the extraction progress as a percentage
            $progressCounter++

            # Extract the ZIP file
            $extractPath = Join-Path -Path $outputFolder -ChildPath $zipFile.BaseName
            try{
                if (-not (Test-Path -Path $extractPath)) {
                    if ($systemIOAvail) {
                        Extract-ZIPFilesystemIO -zipFile $zipFile -extractPath $extractPath
                        Write-ProgressBar -actionText 'Extracting'-zipFileName $zipFile.Name -progressCounter $progressCounter -totalArchives $totalArchives
                    } else {
                        Extract-ZIPFileShellApp -zipFile $zipFile -extractPath $extractPath -shellApp $shellApp
                        Write-ProgressBar -actionText 'Extracting (Fallback Shell)' -zipFileName $zipFile.Name -progressCounter $progressCounter -totalArchives $totalArchives
                    }
                } else {
                    if ($applyActionToFuture -notmatch 'Y') {
                        do {
                            $action = Read-Host "Duplicate found:`n'$($zipFile.Name)' already exists in the output folder.`nChoose an action (O)verwrite, (S)kip, (R)ename (by appending copy #)"
                        } while ($action -notin 'O', 'S', 'R')

                        if ($applyActionToFuture -notmatch 'D'){
                            do {
                                # Ask the user if they want to apply this action to all future duplicates
                                $applyActionToFuture = Read-Host "Apply this action to all future duplicates?`n(Y)es, (N)o, (D)No Don't Ask Again)"
                            } while ($applyActionToFuture -notin 'Y', 'N', 'D')
                        }
                    }

                    if ($applyActionToFuture -in 'Y', 'D' -and $action) {
                        switch ($action.ToUpper()) {
                            'O' {
                                Get-ChildItem $extractPath | Remove-Item -Recurse -Force
                                if ($systemIOAvail) {
                                    Extract-ZIPFilesystemIO -zipFile $zipFile -extractPath $extractPath
                                    Write-ProgressBar -actionText 'Extracting & Overwriting' -zipFileName $zipFile.Name -progressCounter $progressCounter -totalArchives $totalArchives
                                } else {
                                    Extract-ZIPFileShellApp -zipFile $zipFile -extractPath $extractPath -shellApp $shellApp
                                    Write-ProgressBar -actionText 'Extracting & Overwriting (Fallback Shell)' -zipFileName $zipFile.Name -progressCounter $progressCounter -totalArchives $totalArchives
                                }
                            }
                            'S' {
                                Write-ProgressBar -actionText 'Skipping Extraction' -zipFileName $zipFile.Name -progressCounter $progressCounter -totalArchives $totalArchives
                            }
                            'R' {
                                $copyNumber = 1
                                while (Test-Path -Path $extractPath) {
                                    $copyNumber++
                                    $extractPath = Join-Path -Path $outputFolder -ChildPath "$($zipFile.BaseName) (Copy $copyNumber)"
                                }
                                if ($systemIOAvail) {
                                    Extract-ZIPFilesystemIO -zipFile $zipFile -extractPath $extractPath
                                    Write-ProgressBar -actionText 'Extracting & Renaming' -zipFileName $zipFile.Name -progressCounter $progressCounter -totalArchives $totalArchives
                                } else {
                                    Extract-ZIPFileShellApp -zipFile $zipFile -extractPath $extractPath -shellApp $shellApp
                                    Write-ProgressBar -actionText 'Extracting & Renaming (Fallback Shell)' -zipFileName $zipFile.Name -progressCounter $progressCounter -totalArchives $totalArchives
                                }
                            }
                        }
                        if ($applyActionToFuture -eq 'D') {
                            $action = $null
                        }
                    }
                }
            } catch {
                Write-Warning "Failed to process archive. $_"
                continue
            }
        }
    } catch {
        Write-Warning "Failed to extract $($zipFile.Name). $_"
    } finally {
        # Release the Shell.Application COM object when done
        if ($shellApp -ne $null) {
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($shellApp) | Out-Null
        }
    }
}

function Start-Program {
    # Prompt the user for the source folder and set default as .\
    $defaultSourceFolder = ".\"
    $sourceFolder = Read-Host "Enter the source folder path (default: $defaultSourceFolder)"
    if ([string]::IsNullOrEmpty($sourceFolder)) {
        $sourceFolder = $defaultSourceFolder
    }
    if (Test-Path -Path $sourceFolder -PathType Container) {
        # Convert the $sourceFolder to an absolute path
        $absoluteSourceFolder = Convert-Path $sourceFolder
    } else {
        Write-Warning "Unable to verify source path!"
        return    
    }

    # Prompt the user for the destination folder and set default as .\extracted
    $defaultOutputFolder = ".\extracted"
    $outputFolder = Read-Host "Enter the destination folder path (default: $defaultOutputFolder)"
    if ([string]::IsNullOrEmpty($outputFolder)) {
        $outputFolder = $defaultOutputFolder
    }
    # Verify that the source folder exists before proceeding
    if (-not (Test-Path -Path $sourceFolder -PathType Container)) {
        Write-Warning "Source folder '$sourceFolder' not found."
        return
    }

    # Verify that the output folder exists or create it if it doesn't
    if (-not (Test-Path -Path $outputFolder -PathType Container)) {
        New-Item -ItemType Directory -Path $outputFolder -ErrorAction Stop | Out-Null
    }

    if (Test-Path -Path $outputFolder -PathType Container) {
        # Convert the $outputFolder to an absolute path
        $absoluteOutputFolder = Convert-Path $outputFolder
    } else {
        Write-Warning "Unable to verify destination path!"
        $absoluteOutputFolder = Convert-Path $outputFolder
        return
    }

    Extract-ZIPFiles -sourceFolder $absoluteSourceFolder -outputFolder $absoluteOutputFolder
}

Start-Program