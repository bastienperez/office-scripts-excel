// Script created by B.Perez 20240719
// This script converts CSV data in each row of the active sheet into columns

function main(workbook: ExcelScript.Workbook) {
  // Get the active sheet
  let worksheet = workbook.getActiveWorksheet();

  // Get all values from the active sheet as a 2D array
  let range = worksheet.getUsedRange();
  let values = range.getValues() as string[][];

  // Iterate through each row to convert CSV data into columns
  for (let row = 0; row < values.length; row++) {
    let csvLine = values[row][0] as string;
    let csvValues = csvLine.split(','); // Change ',' to your CSV delimiter if necessary

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
  
  // Automatically adjusts the width of all columns in the worksheet
  // to fit the content. The selected range spans from column A to XFD (the last available column).
  worksheet.getRange("A:XFD").getFormat().autofitColumns();

  // Automatically adjusts the height of all rows in the worksheet
  // to fit the content. The selected range covers all rows and columns.
  worksheet.getRange("A:XFD").getFormat().autofitRows;
}