const XLSX = require('xlsx');
const workbook = XLSX.readFile('LINE OA 加入資料_分析結果.xlsx');

console.log("Sheet names:", workbook.SheetNames);

const sheet1 = workbook.Sheets[workbook.SheetNames[0]];
const data1 = XLSX.utils.sheet_to_json(sheet1, { range: 1 });
const totalJoinsSum = data1.reduce((s, r) => s + (Number(r['總加入人數']) || 0), 0);
const lastRowJoins = data1[data1.length - 1]['總加入人數'];
console.log("Sheet 1 Sum (總加入人數):", totalJoinsSum);
console.log("Sheet 1 Last Row (總加入人數):", lastRowJoins);

const sheet2 = workbook.Sheets['Data_回覆時程'];
const data2 = XLSX.utils.sheet_to_json(sheet2);
const rowCount = data2.length;
const solvedCount = data2.filter(r => r['已回覆'] == 1).length;
const totalDays = data2.reduce((s, r) => s + (Number(r['回覆天數']) || 0), 0);

console.log("Sheet 2 Rows:", rowCount);
console.log("Sheet 2 Solved Count:", solvedCount);
console.log("Sheet 2 Overall Resp Rate:", (solvedCount / rowCount) * 100);
console.log("Sheet 2 Avg Resp Days:", totalDays / rowCount);

// Check if there are other columns that might affect calculations
console.log("Sheet 2 Sample row properties:", Object.keys(data2[0]));
