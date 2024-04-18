const ExcelJS = require('exceljs');
const fs = require('fs');

const tmpWorkBook = new ExcelJS.Workbook();
const outputPath = 'new_img.xlsx'; // Output file path for the new Excel file

const worksheet = tmpWorkBook.addWorksheet('Sheet1');
var imageID = tmpWorkBook.addImage({
    filename: 'hello 1.png',
    extension: 'png',
});
console.log("imageID", imageID);
// Define the cell where you want to add the image
const cell = worksheet.getCell('A1');
const relCol = cell._column.number - 1; // Adjust for zero-based indexing
const relNum = cell.row - 1; // Adjust for zero-based indexing

worksheet.addImage(imageID, {
    tl: { col: relCol, row: relNum },
    ext: { width: 100, height: 50 } // currently I'm using fixed width and height.
});

tmpWorkBook.xlsx.writeFile(outputPath)
    .then(function() {
        console.log("Image added successfully to the new Excel file:", outputPath);
    })
    .catch(function(error) {
        console.error("Error:", error);
    });
