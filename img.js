const ExcelJS = require('exceljs');

const tmpWorkBook = new ExcelJS.Workbook();
const outputPath = 'new_img1.xlsx'; // Output file path for the new Excel file

const worksheet = tmpWorkBook.addWorksheet('Sheet1');
var imageID = tmpWorkBook.addImage({
    filename: 'hello 1.png',
    extension: 'png',
});
console.log("imageID", imageID);

// Define the cell where you want to add the image (for example, B5)
const targetCell = worksheet.getCell('I6');

worksheet.addImage(imageID, {
    tl: { col: targetCell._column.number - 1, row: targetCell.row - 1 }, // Adjust for zero-based indexing
    ext: { width: 100, height: 50 } // currently I'm using fixed width and height.
});

tmpWorkBook.xlsx.writeFile(outputPath)
    .then(function() {
        console.log("Image added successfully to the new Excel file:", outputPath);
    })
    .catch(function(error) {
        console.error("Error:", error);
    });
