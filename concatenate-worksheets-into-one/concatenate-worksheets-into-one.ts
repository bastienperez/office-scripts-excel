// Script created by B.Perez 20241023

function main(workbook: ExcelScript.Workbook) {
  // Get all worksheets in the workbook
  let sheets = workbook.getWorksheets();

  // Create a new worksheet for the combined data
  let newSheet = workbook.addWorksheet("CombinedData");

  // Initialize a variable to track the current row for pasting data
  let currentRow = 0;

  // Get the header from the first sheet
  let firstSheet = sheets[0];
  let header = firstSheet.getUsedRange().getValues()[0];

  // Add the "Source" column header to the header array
  let updatedHeader = [...header, "Source"];

  // Set the header in the new sheet
  newSheet.getRangeByIndexes(currentRow, 0, 1, updatedHeader.length).setValues([updatedHeader]);

  // Move to the next row after the header
  currentRow++;

  // Loop through each sheet
  sheets.forEach((sheet, index) => {
    // Get the used range of the current sheet
    let usedRange = sheet.getUsedRange();

    // Skip the header row for all sheets
    let startRow = 1;

    // Get the data from the used range, excluding the header
    let data = usedRange.getValues().slice(startRow);

    // Add sheet name as the last column for each row
    let dataWithSheetName = data.map(row => {
      row.push(sheet.getName());
      return row;
    });

    // Paste the data into the new sheet starting at the current row
    newSheet.getRangeByIndexes(currentRow, 0, dataWithSheetName.length, dataWithSheetName[0].length).setValues(dataWithSheetName);

    // Update the current row for the next paste operation
    currentRow += dataWithSheetName.length;
  });

  // Auto-fit columns in the new sheet
  newSheet.getUsedRange().getFormat().autofitColumns();

  // Convert the used range to a table
  let lastUsedRange = newSheet.getUsedRange();
  let table = newSheet.addTable(lastUsedRange, true);
  table.setName("CombinedDataTable");

  // Apply default table style (e.g., "TableStyleMedium9")
  table.setPredefinedTableStyle("TableStyleMedium2");
}