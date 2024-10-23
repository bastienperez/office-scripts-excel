# Concatenate All Worksheets into One

## Description

This script combines data from all worksheets in an Excel workbook into a new worksheet called `CombinedData`. It merges all data while avoiding duplicate headers and adds a new column that contains the name of the source worksheet for each row.

## Features

- **Sheet Consolidation**: Combines data from multiple worksheets into a single sheet.
- **Avoid Header Duplication**: Copies headers from only the first worksheet.
- **Source Column**: Adds a column named "source" to track the worksheet of origin for each row.
- **Efficient Data Processing**: Automatically excludes headers from subsequent sheets, ensuring clean data concatenation.

## Requirements

- **Uniform Headers**: All worksheets must have the same headers for the script to function correctly. Any mismatch in headers between sheets will cause the script to fail.

## Usage

1. Ensure all worksheets have identical headers.
2. Open your workbook in Excel.
3. Run the provided Office Script to generate a new worksheet with the concatenated data.
4. The "source" column will indicate the original sheet for each row.

### Example Execution

Here is an animated image showing the script in action:

![Script Execution](concatenate-worksheets-into-one.gif)