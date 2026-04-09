const XLSX = require('xlsx');
const workbook = XLSX.readFile('LINE OA 加入資料_分析結果.xlsx');
const firstSheetName = workbook.SheetNames[0];
const worksheet = workbook.Sheets[firstSheetName];
const data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

console.log("Headers:");
console.log(data[0]);

console.log("\nSample Rows:");
console.log(data.slice(1, 4));

// If there are multiple sheets, list them
if (workbook.SheetNames.length > 1) {
    console.log("\nOther Sheets:", workbook.SheetNames.slice(1));
}
