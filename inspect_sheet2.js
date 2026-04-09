const XLSX = require('xlsx');
const workbook = XLSX.readFile('LINE OA 加入資料_分析結果.xlsx');
const secondSheetName = 'Data_回覆時程';
const worksheet = workbook.Sheets[secondSheetName];
const data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

console.log("Headers for " + secondSheetName + ":");
console.log(data[0]);

console.log("\nSample Rows:");
console.log(data.slice(1, 4));
