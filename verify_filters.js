const XLSX = require('xlsx');
const workbook = XLSX.readFile('LINE OA 加入資料_分析結果.xlsx');

const sheet2 = workbook.Sheets['Data_回覆時程'];
const data2 = XLSX.utils.sheet_to_json(sheet2);

const filtered2 = data2.filter(r => r['非測試'] === '非測試');
console.log("Filtered Sheet 2 Rows (非測試):", filtered2.length);

const solvedCount = filtered2.filter(r => r['已回覆'] == 1).length;
const totalDays = filtered2.reduce((s, r) => s + (Number(r['回覆天數']) || 0), 0);

console.log("Filtered Sheet 2 Solved Count:", solvedCount);
console.log("Filtered Sheet 2 Overall Resp Rate:", (solvedCount / filtered2.length) * 100);
console.log("Filtered Sheet 2 Avg Resp Days:", totalDays / filtered2.length);

// Check Sheet 1 logic again for 2074
const sheet1 = workbook.Sheets[workbook.SheetNames[0]];
const data1 = XLSX.utils.sheet_to_json(sheet1, { range: 1 });
// Let's print the last few rows of Sheet 1
console.log("Sheet 1 Last 5 rows:", data1.slice(-5));
