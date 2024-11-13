// Script created by B.Perez
// v1.0 - 2024-07-19 - Initial version
// v1.1 - 2024-11-13 - Added auto-filter to header row and fix column with comma

// This script converts CSV data in each row of the active sheet into columns
// Automatically adjusts the width of all columns in the worksheet to fit the content
// Automatically adjusts the height of all rows in the worksheet to fit the content
// Adds an auto-filter to the header row (assumed to be the first row)

function main(workbook: ExcelScript.Workbook) {
  // Get the active sheet
  let worksheet = workbook.getActiveWorksheet();

  // Get all values from the active sheet as a 2D array
  let range = worksheet.getUsedRange();
  let values = range.getValues() as string[][];

  // Iterate through each row to convert CSV data into columns
  for (let row = 0; row < values.length; row++) {
    let csvLine = values[row][0] as string;

    // Parse CSV using regex to handle values with commas inside quotes
    let csvValues = csvLine.match(/(".*?"|[^",\s]+)(?=\s*,|\s*$)/g) || []; // Matches quoted values or unquoted values

    // If the row contains CSV data
    if (csvValues.length > 1) {
      // Write the CSV values into the same row
      for (let col = 0; col < csvValues.length; col++) {
        // Remove leading and trailing quotes from each CSV value
        let cleanedValue = csvValues[col].replace(/^"(.*)"$/, '$1');
        worksheet.getCell(row, col).setValue(cleanedValue);
      }
    }
  }

  // Automatically adjusts the width of all columns in the worksheet to fit the content
  worksheet.getRange("A:XFD").getFormat().autofitColumns();

  // Automatically adjusts the height of all rows in the worksheet to fit the content
  worksheet.getRange("A:XFD").getFormat().autofitRows();

  // Define the range for the header row (assumed to be the first row) and apply an auto-filter
  let headerRange = worksheet.getRange("A1:" + worksheet.getCell(0, values[0].length - 1).getAddress());
  worksheet.getAutoFilter().apply(headerRange);
}