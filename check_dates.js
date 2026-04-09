const XLSX = require('xlsx');
const workbook = XLSX.readFile('LINE OA 加入資料_分析結果.xlsx');

const sheet1 = workbook.Sheets['各管道每日加入人數'];
const data1 = XLSX.utils.sheet_to_json(sheet1, { range: 1 });

const dailyJoins = data1.filter(r => typeof r['加入LINE OA日期'] === 'number');
const dates = dailyJoins.map(r => XLSX.utils.format_cell({ v: r['加入LINE OA日期'], t: 'd' }));
console.log("Date Range in Sheet 1:", dates[0], "to", dates[dates.length - 1]);

const sheet2 = workbook.Sheets['Data_回覆時程'];
const data2 = XLSX.utils.sheet_to_json(sheet2);
const dates2 = data2.map(r => r['預約時間']).filter(d => d).sort();
console.log("Date Range in Sheet 2:", dates2[0], "to", dates2[dates2.length - 1]);
