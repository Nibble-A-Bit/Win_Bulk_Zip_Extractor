# Win_Mass_Extractor - PowerShell Script to Simplify Extracting Multiple ZIP Archives

![Win_Mass_Extractor](https://link.to.your.image)

## Overview

Win_Mass_Extractor is a PowerShell script designed to simplify the process of extracting multiple ZIP archives in a given folder and its subfolders. The script leverages the [System.IO.Compression.ZipFile] class available in PowerShell 5.0 or higher to perform the extraction. If the required class is not available, the script falls back to using the Shell.Application COM object to achieve the same functionality.

## Features

- Extracts multiple ZIP archives in a folder and its subfolders.
- Utilizes [System.IO.Compression.ZipFile] for faster extraction (PowerShell 5.0 or higher).
- Falls back to Shell.Application COM object if [System.IO.Compression.ZipFile] is not available.
- Provides user interaction options for handling duplicate ZIP files.
- Displays a progress bar during the extraction process.

## Prerequisites

- PowerShell 5.0 or higher.

## How to Use

1. Clone or download the repository to your local machine.

2. Open a PowerShell terminal.

3. Change the directory to the location of the script file.

4. Run the script using the following command:

   ```
   .\Win_Mass_Extractor.ps1
   ```

5. The script will prompt you for the source folder path. Enter the path to the folder containing the ZIP archives you want to extract. If no input is provided, the default source folder is set to the current directory (.\).

6. Next, the script will prompt you for the destination folder path. Enter the path where you want the extracted files to be stored. If no input is provided, the default destination folder is set to .\extracted.

7. The script will then process all the ZIP archives found in the source folder and its subfolders. It will extract the contents to the specified destination folder.

8. If the script encounters duplicate ZIP files in the destination folder, it will prompt you to choose an action:
   - O: Overwrite the existing files.
   - S: Skip the extraction for the current ZIP archive.
   - R: Rename the current ZIP archive before extraction.

   You can also choose to apply the selected action to all future duplicates.

## Progress Bar

During the extraction process, the script displays a progress bar indicating the progress of the extraction. The progress bar shows the percentage of completion for the entire extraction operation.

## Example

```powershell
PS C:\Projects\Win_Mass_Extractor> .\Win_Mass_Extractor.ps1
Enter the source folder path (default: .\): C:\Projects\ZIP_Archives
Enter the destination folder path (default: .\extracted): C:\Projects\Extracted_Files
Confirmed PowerShell 5.V
[Duplicating Files      ] archive1.zip                                      25.00% [====      ] 1/4
[Extracting             ] archive2.zip                                      50.00% [========  ] 2/4
[Skipping Extraction    ] archive3.zip                                      75.00% [==========] 3/4
[Extracting & Overwriting (Fallback Shell)] archive4.zip                    100.00% [==========] 4/4
```

## License

This project is licensed under the [MIT License](https://opensource.org/licenses/MIT).

## Contributions

Contributions to the project are welcome. If you find any issues or have suggestions for improvements, feel free to create a pull request or submit an issue.

## Disclaimer

This script is provided as-is, without any warranty or guarantee of any kind. The author is not responsible for any damage caused by the usage of this script.

## Credits

This script was created by https://github.com/KR34T1V and inspired by https://chat.openai.com/.

---

Thank you for using Win_Mass_Extractor! If you have any questions or feedback, please don't hesitate to reach out. Happy extracting!