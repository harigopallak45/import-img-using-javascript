const ExcelJS = require('exceljs');

// Load input Excel file
const workbook = new ExcelJS.Workbook();
workbook.xlsx.readFile('input data.xlsx')
    .then(() => {
        // Get worksheet
        const worksheet = workbook.getWorksheet(1); // Assuming the data is in the first worksheet

        // Define column indices
        const column1Index = 1; // Column A
        const column2Index = 2; // Column B
        const column3Index = 3; // Column C

        // Operation 1: Double values in column 2 based on column 1
        worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
            const value1 = row.getCell(column1Index).value;
            if (value1) {
                const newValue2 = value1 * 2;
                row.getCell(column2Index).value = newValue2;
            }
        });

        // Operation 2: Sum values in column 1 and 2 and write to column 3
        worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
            const value1 = row.getCell(column1Index).value;
            const value2 = row.getCell(column2Index).value;
            if (value1 && value2) {
                const sum = value1 + value2;
                row.getCell(column3Index).value = sum;
            }
        });

        // Write output Excel file
        return workbook.xlsx.writeFile('output data.xlsx');
    })
    .then(() => {
        console.log('Operations completed successfully!');
    })
    .catch(err => {
        console.error('Error:', err);
    });
