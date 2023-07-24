# Win Mass Extractor

Win Mass Extractor is a PowerShell script that simplifies the process of extracting multiple ZIP archives to a single folder recursively. With this script, you can quickly and efficiently extract the contents of all ZIP files within a specified source folder and its subfolders into a designated output folder.

## How to Use

1. **Clone the Repository**: Clone this repository to your local machine.

2. **Run the Script**: Open a PowerShell window and navigate to the folder where you cloned the repository. Run the script by executing the `ExtractArchives.ps1` script with the appropriate parameters.

   ```powershell
   .\ExtractArchives.ps1 -SourceFolder "C:\Path\To\Source\Folder" -OutputFolder "C:\Path\To\Output\Folder"
   ```

   Replace `"C:\Path\To\Source\Folder"` with the path to the folder containing the ZIP archives you want to extract, and `"C:\Path\To\Output\Folder"` with the path to the output folder where you want to store the extracted contents.

3. **Sit Back and Relax**: The script will automatically locate all ZIP files within the source folder and its subfolders. It will then extract their contents, maintaining the folder structure, and merge them into the specified output folder.

## Features

- Recursively Extract ZIP Archives: The script traverses through the source folder and all its subfolders, locating and extracting ZIP files to a single output folder.

- Preserve Folder Structure: The extracted files retain their original folder structure to ensure that you can easily find and organize the contents.

- Efficient and Time-Saving: With Win Mass Extractor, you no longer need to extract each ZIP file manually. The process is automated, saving you time and effort.

## Requirements

- Windows OS: The script is designed to work on Windows operating systems with PowerShell support.

## Limitations

- Only ZIP Archives Supported: Currently, the script supports extracting only ZIP archives. Other archive formats are not supported.

## License

This project is licensed under the [MIT License](LICENSE).

## Contribution

If you encounter any issues, have suggestions, or want to contribute to this project, feel free to open an issue or submit a pull request. Your contributions are highly appreciated!

## Disclaimer

While the script is designed to perform its function efficiently, it's essential to exercise caution when using it. Ensure that you have appropriate permissions and backups in place before running the script.

**USE AT YOUR OWN RISK!**

---

Thank you for using Win Mass Extractor! We hope this script simplifies your archive extraction process and enhances your productivity. If you have any questions or feedback, please don't hesitate to reach out. Happy extracting!
