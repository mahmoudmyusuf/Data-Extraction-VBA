| \[03/02/2022\] |
|----------------|

| Data-Extraction-VBA |
|--------------------------------|

 Mahmoud Mohamed Abdel Aziz
 
| <img src="media/VBA-Extract.jpg" alt="Project Image" width="600" style="float: right; margin-left: 15px; margin-bottom: 15px;" /> |
|:--:|

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


## Steps

1. **Set Up the Workbook:**
   - Open the workbook in Excel and enable macros.
   - Ensure you have a source file from which data will be pulled. 
   - Copy and paste the links to all the required data from the source file into the destination file. Once done, the code will automate the process of updating the links.
   - Enter the folder path in cell `T2` in the "Data" sheet.
   - Specify folder name and file name patterns in cells `T3` and `T4`.
   
2. **Confirm a Single Link:**
   - In Excel, navigate to **File > Edit Links to Files** to check the linked files.
   - Confirm that **only one link** exists in the list. If there are multiple links, the code will only update the first link in the list.
   
3. **Run the VBA Code:**
   - When you run the code, it will automatically process only new or edited files, updating the links to match the latest source data file. It avoids reprocessing files that have already been handled.

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

**Key Performance Features**:
- It processes only new or edited files, avoiding the repetition of previously processed files. This is done by checking each file's path and timestamp, ensuring that only new files are added to the workbook.
- It handles data from all sheets in the source files as it relies on updating links, rather than updating each cell individually. This approach ensures that all data is synchronized efficiently without manually processing each worksheet.


## Conclusion

This VBA project is a powerful automation tool for handling report files and updating workbooks. It is customizable, allowing you to define the folder and file name patterns for each use case. The error handling ensures a smooth experience even when working with large datasets.
