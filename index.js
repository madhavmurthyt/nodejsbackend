const fs = require('fs');
const XLSX = require('xlsx');
const workbook = XLSX.readFile('Book1.xlsx');
const worksheet = workbook.Sheets['Sheet1']; 

console.log(worksheet['A1'].v);