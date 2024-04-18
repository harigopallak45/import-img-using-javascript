const XLSX = require('xlsx');

// Sample JSON data
const jsonData = [
  { "Title 1": "Value 1", "Title 2": 10, "Title 3": true },
  { "Title 1": "Value 2", "Title 2": 20, "Title 3": false },
  { "Title 1": "Value 3", "Title 2": 30, "Title 3": true }
];

// Create a new workbook
const workbook = XLSX.utils.book_new();

// Create a worksheet
const worksheet = XLSX.utils.json_to_sheet(jsonData);

// Add worksheet to workbook
XLSX.utils.book_append_sheet(workbook, worksheet, "JSON Data");

// Write workbook to a file
XLSX.writeFile(workbook, 'output_data.xlsx');

console.log('JSON data successfully written to Excel file.');
