// const fs = require('fs');
// const XLSX = require('xlsx');
// const workbook = XLSX.readFile('Book1.xlsx');
// const worksheet = workbook.Sheets['Sheet1']; 

// console.log(worksheet['A1'].v);

const fs = require('fs');
const xlsx = require('xlsx');

// Function to get the size of a folder recursively
function getFolderSize(path) {
    let totalSize = 0;
    const files = fs.readdirSync(path);
    files.forEach(file => {
        const filePath = `${path}/${file}`;
        console.log(filePath);
        const stats = fs.statSync(filePath);
        if (stats.isDirectory()) {
            totalSize += getFolderSize(filePath);
        } else {
            totalSize += stats.size;
        }
    });
    console.log("Total Size "+totalSize);
    return totalSize;
}

// Load the Excel file
const workbook = xlsx.readFile('Book1.xlsx');
const sheetName = workbook.SheetNames[0];
const worksheet = workbook.Sheets[sheetName];

// Get the range of cells with folder paths (assuming they are in column A)
const range = worksheet['!ref'];
console.log("range "+range);
const cells = xlsx.utils.decode_range(range);
const folderPaths = [];
for (let rowNum = cells.s.r; rowNum <= cells.e.r; rowNum++) {
    const cellAddress = xlsx.utils.encode_cell({ r: rowNum, c: 0 }); // Column A
    const folderPath = worksheet[cellAddress]?.v;
    console.log("folderPath "+folderPath);
    if (folderPath) {
        folderPaths.push(folderPath);
    }
}

// Calculate folder sizes and write them to the spreadsheet
folderPaths.forEach((folderPath, index) => {
    const size = getFolderSize(folderPath);
    console.log("size "+size);
    const cellAddress = xlsx.utils.encode_cell({ r: index, c: 1 }); // Column B
    console.log("cellAddress "+cellAddress);
    worksheet[cellAddress] = { v: size, t: 'n' };
});

// Save the updated workbook
xlsx.writeFile(workbook, 'Book1.xlsx');