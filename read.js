const xlsx = require('xlsx');

// Path to the input Excel file
const inputFilePath = 'AJJAMPURA_110_15_10_2023.xls';

// Read the Excel file
const workbook = xlsx.readFile(inputFilePath);
const sheetName = workbook.SheetNames[0];
const worksheet = workbook.Sheets[sheetName];

// Function to extract all string values from a specified row, excluding values containing "DUMMY"
function extractRowValues(worksheet, rowNumber) {
  const rowValues = [];
  const range = xlsx.utils.decode_range(worksheet['!ref']);

  for (let col = range.s.c; col <= range.e.c; col++) {
    const cellAddress = xlsx.utils.encode_cell({ c: col, r: rowNumber });
    const cell = worksheet[cellAddress];
    const value = cell ? cell.v : null; // Use null for empty cells

    // Check if the value is a string and does not contain "DUMMY"
    if (value !== null && typeof value === 'string' && !value.includes('DUMMY')) {
      rowValues.push({ value, col });
    }
  }

  return rowValues;
}

// Function to extract data from a specific column starting from a given row for a certain number of rows
function extractColumnData(worksheet, col, startRow, numRows) {
  const columnData = [];
  for (let row = startRow; row < startRow + numRows; row++) {
    const cellAddress = xlsx.utils.encode_cell({ c: col, r: row });
    const cell = worksheet[cellAddress];
    const value = cell ? cell.v : 'N/A'; // Use 'N/A' for empty cells
    columnData.push(value);
  }
  return columnData;
}

// Extract string values from the 3rd row (row index 2, as it's 0-based)
const thirdRowValues = extractRowValues(worksheet, 2);

// Log the extracted values
console.log('String values in the third row (excluding values containing "DUMMY"):', thirdRowValues);

// Extract string values from the 2003rd row (row index 2002, as it's 0-based)
const twoThousandThirdRowValues = extractRowValues(worksheet, 2002);

// Log the extracted values
console.log('String values in the 2003rd row (excluding values containing "DUMMY"):', twoThousandThirdRowValues);

// Extract string values from the 4003rd row (row index 4002, as it's 0-based)
const fourThousandThirdRowValues = extractRowValues(worksheet, 4002);

// Log the extracted values
console.log('String values in the 4003rd row (excluding values containing "DUMMY"):', fourThousandThirdRowValues);

// Extract string values from the 6003rd row (row index 6002, as it's 0-based)
const sixThousandThirdRowValues = extractRowValues(worksheet, 6002);

// Log the extracted values
console.log('String values in the 6003rd row (excluding values containing "DUMMY"):', sixThousandThirdRowValues);

// Prepare data for the new worksheet
const outputData = [['SCADA Tag', 'DATA', 'Range']]; // Include "Range" column in the header

// Extract data for each SCADA tag from the 3rd row and add to the outputData
thirdRowValues.forEach(({ value, col }) => {
  const columnData = extractColumnData(worksheet, col, 5, 1440); // Extract data from 6th row (index 5) and next 1440 rows
  columnData.forEach(dataValue => {
    outputData.push([value, dataValue, '0-2000']); // Include "Range" value for each row
  });
});

// Add 4 empty rows as a gap between the tags from the 3rd row and the 2003rd row
for (let i = 0; i < 4; i++) {
  outputData.push([]);
}

// Extract data for each SCADA tag from the 2003rd row and add to the outputData
twoThousandThirdRowValues.forEach(({ value, col }) => {
  const columnData = extractColumnData(worksheet, col, 2006, 1440); // Extract data from 2007th row (index 2006) and next 1440 rows
  columnData.forEach(dataValue => {
    outputData.push([value, dataValue, '2000-4000']); // Include "Range" value for each row
  });
});

// Add 4 empty rows as a gap between the tags from the 2003rd row and the 4003rd row
for (let i = 0; i < 4; i++) {
  outputData.push([]);
}

// Extract data for each SCADA tag from the 4003rd row and add to the outputData
fourThousandThirdRowValues.forEach(({ value, col }) => {
  const columnData = extractColumnData(worksheet, col, 4006, 1440); // Extract data from 4007th row (index 4006) and next 1440 rows
  columnData.forEach(dataValue => {
    outputData.push([value, dataValue, '4000-6000']); // Include "Range" value for each row
  });
});

// Add 4 empty rows as a gap between the tags from the 4003rd row and the 6003rd row
for (let i = 0; i < 4; i++) {
  outputData.push([]);
}

// Extract data for each SCADA tag from the 6003rd row and add to the outputData
sixThousandThirdRowValues.forEach(({ value, col }) => {
  const columnData = extractColumnData(worksheet, col, 6006, 1440); // Extract data from 6007th row (index 6006) and next 1440 rows
  columnData.forEach(dataValue => {
    outputData.push([value, dataValue, '6000-8000']); // Include "Range" value for each row
  });
});

// Create a new workbook and worksheet for the output data
const outputWorkbook = xlsx.utils.book_new();
const outputWorksheet = xlsx.utils.aoa_to_sheet(outputData);
xlsx.utils.book_append_sheet(outputWorkbook, outputWorksheet, 'Sheet1');

// Save the output workbook to an Excel file
const outputFilePath = 'Extracted_SCADA_Tag_Data.xlsx';
xlsx.writeFile(outputWorkbook, outputFilePath);

console.log('SCADA Tag data has been successfully written to Extracted_SCADA_Tag_Data.xlsx');
