| \[03/02/2022\] |
|----------------|

| Data-Extraction-VBA |
|--------------------------------|

 Mahmoud Mohamed Abdel Aziz


## Overview

This project is designed to automate the collection and processing of report files using Excel VBA. The main goals of the project are to:
- Scan a directory for new files.
- Extract data from the files and update a master workbook.
- Handle potential errors in file retrieval and processing.

## Code Breakdown

### `Get_rest_file`:
This macro scans a folder and its subfolders to locate report files based on the pattern specified in the workbook. It stores the file paths of the new reports in an array (`NewFiles`) for later processing.

**Key features**:
- Scans subfolders recursively.
- Uses patterns from the `Data` sheet to match file and folder names.
- Avoids duplicates by checking if a file is already added.

### `Update_WB_link`:
This macro updates the workbook with data from the new files that were found by the `Get_rest_file` macro. It updates links, file paths, and processes data into the workbook's "Data" sheet.

**Key features**:
- Updates file paths and links in the workbook.
- Performs a Find and Replace operation to clean up data (e.g., replace "N/A" with an empty string).
- Saves the workbook and logs the time taken for the operation.

## How to Use

### Setup
1. Open the workbook in Excel and enable macros.
2. Enter the folder path in cell `T2` in the "Data" sheet.
3. Specify folder name and file name patterns in cells `T3` and `T4`.

### Running the Macros
1. Press `Alt + F8` to open the "Macro" dialog.
2. Select the `Update_WB_link` macro and click "Run."
3. The script will process all the new report files, update the workbook, and display the results.

## Error Handling

The code includes error handling to ensure that if there are issues with the file paths or if any unexpected errors occur, they are properly addressed:
- If the folder path is incorrect, a message box will appear to notify the user.
- The program ensures that no duplicate files are added to the "Data" sheet by checking file paths.

## Performance
The program is designed to handle a large number of files and folders efficiently. It uses arrays to store file paths temporarily, which allows for quick processing.

## Conclusion

This VBA project is a powerful automation tool for handling report files and updating workbooks. It is customizable, allowing you to define the folder and file name patterns for each use case. The error handling ensures a smooth experience even when working with large datasets.
