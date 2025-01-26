# Excel VBA Report Project

This project automates the process of gathering and processing report files using Excel VBA. It scans a folder and its subfolders for specific report files, extracts data from them, and updates an Excel workbook accordingly.

![Project Image](Documentation/ERD.png)

## Table of Contents
- [Project Overview](#project-overview)
- [Features](#features)
- [How to Use](#how-to-use)
- [How to Contribute](#how-to-contribute)
- [License](#license)

## Project Overview

This project is designed to automate the process of loading report files from various directories and then updating an Excel workbook. The two main VBA functions are:
- **Get_rest_file**: Scans the specified folder and subfolders for new report files.
- **Update_WB_link**: Updates links in the Excel workbook with data from the new files.

## Features

- Recursively scans folders and subfolders for specific files.
- Retrieves file data and updates the main workbook.
- Updates external file links in the workbook.
- Automatic error handling for missing files and data.
- Updates the workbook’s "Data" sheet with the latest file paths.

## How to Use

1. **Clone the repository** to your local machine:
   ```bash
   git clone https://github.com/yourusername/Excel-VBA-Report-Project.git
