// Script created by B.Perez
// v1.0 - 2024-07-19 - Initial version
// v1.1 - 2024-11-13 - Added auto-filter to header row and fix column with comma
// v1.2 - 2024-12-06 - Fix issue with column

// This script converts CSV data in each row of the active sheet into columns
// Automatically adjusts the width of all columns in the worksheet to fit the content
// Automatically adjusts the height of all rows in the worksheet to fit the content
// Adds an auto-filter to the header row (assumed to be the first row)

function main(workbook: ExcelScript.Workbook) {
	let worksheet = workbook.getActiveWorksheet();
	let range = worksheet.getUsedRange();
	let values = range.getValues() as string[][];

	for (let row = 0; row < values.length; row++) {
		let csvLine = values[row][0] as string;

		// Parse CSV with a regex that correctly handles empty fields, quoted strings, and commas
		let csvValues =
			csvLine.match(/(?<=^|,)(?:"([^"]*(?:""[^"]*)*)"|[^,]*)/g) || [];

		if (csvValues.length > 1) {
			for (let col = 0; col < csvValues.length; col++) {
				// Handle quoted strings and unescape double quotes
				let cleanedValue = csvValues[col]
					.replace(/^"(.*)"$/, "$1") // Remove enclosing quotes
					.replace(/""/g, '"'); // Unescape double quotes
				worksheet.getCell(row, col).setValue(cleanedValue || ""); // Handle empty fields
			}
		}
	}

	// Adjust column widths and row heights
	worksheet.getUsedRange().getFormat().autofitColumns();
	worksheet.getUsedRange().getFormat().autofitRows();

	// Apply autofilter to the header row
	let headerRow = worksheet.getRange(
		"A1:" + worksheet.getCell(0, values[0].length - 1).getAddress()
	);
	worksheet.getAutoFilter().apply(headerRow);
}