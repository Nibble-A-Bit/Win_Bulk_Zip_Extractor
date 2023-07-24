# Win_Mass_Extractor

Win_Mass_Extractor is a PowerShell script that simplifies the process of extracting multiple ZIP archives to a single folder recursively. With this script, you can quickly and efficiently extract the contents of all ZIP files within a specified source folder and its subfolders into a designated output folder.

## Features

- Recursive extraction of ZIP archives from a source folder and its subfolders.
- Overwrite prompt with options to overwrite, overwrite all, skip, or rename files (by appending copy #).
- Text-based progress bar displaying extraction progress for each archive.
- Ability to handle duplicate files with customizable actions.
- Automatic creation of the output folder if it does not exist.

## Requirements

- Windows operating system.
- PowerShell version 5.1 or higher.

## Usage

1. Place the Win_Mass_Extractor.ps1 script in a convenient location on your system.

2. Open a PowerShell terminal and navigate to the folder where you placed the script.

3. Run the script by entering the following powershell command (ensure you are in the same directory as the script): .\Win_Mass_Extractor.ps1


4. The script will prompt you for the source directory and destination directory.

5. Optionally, you can press Enter without entering a path to use the default values: `.\` for the source directory and `.\extracted` for the destination directory.

6. The script will then process all ZIP files in the specified source directory and its subfolders, displaying progress updates as each archive is extracted.

7. For any duplicate files encountered during extraction, the script will prompt you to choose an action: overwrite, overwrite all, skip, or rename the file.

## Caveats and Tips

- Ensure you have the necessary permissions to read from the source directory and write to the destination directory.

- Be cautious when using the "Overwrite All" option, as it will overwrite all files without further prompts.

- The script currently supports only ZIP file extraction. If you need to handle other archive formats, you may need to modify the script accordingly.

- Remember that extracting large archives or a large number of archives may take time, depending on your system's performance.

- If you encounter any errors during extraction, ensure that the ZIP files are not corrupted and that you have the required permissions.

- Before running the script on important data, consider testing it on a small subset of files to ensure it behaves as expected.

## Disclaimer

This script is provided as-is, without any warranties or guarantees. Use it at your own risk. The author is not responsible for any data loss or damages caused by using this script.

## License

Win_Mass_Extractor is open-source and distributed under the [MIT License](LICENSE). Feel free to modify and use the script according to your needs.

