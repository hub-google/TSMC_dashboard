const XLSX = require('xlsx');
const workbook = XLSX.readFile('LINE OA 加入資料_分析結果.xlsx');

const sheet1 = workbook.Sheets['各管道每日加入人數'];
const data1 = XLSX.utils.sheet_to_json(sheet1, { range: 1 });

// Filter only rows where the first column is a date (number)
const dailyJoins = data1.filter(r => typeof r['加入LINE OA日期'] === 'number');
const hrSum = dailyJoins.reduce((s, r) => s + (Number(r['由HR公告加入']) || 0), 0);
const friendSum = dailyJoins.reduce((s, r) => s + (Number(r['由好友推薦加入']) || 0), 0);
const cardSum = dailyJoins.reduce((s, r) => s + (Number(r['由服務小卡加入']) || 0), 0);
const totalSumInSheet = dailyJoins.reduce((s, r) => s + (Number(r['總加入人數']) || 0), 0);

console.log("Sheet 1 Real Sums (Daily rows only):");
console.log(`HR Sum: ${hrSum}`);
console.log(`Friend Sum: ${friendSum}`);
console.log(`Card Sum: ${cardSum}`);
console.log(`Total Sum In Sheet: ${totalSumInSheet}`);

const sheet2 = workbook.Sheets['Data_回覆時程'];
const data2 = XLSX.utils.sheet_to_json(sheet2);

// In the image, 175 is the total. Let's see if 175 rows match some criteria.
// Maybe AREA_NM or DEPT_NM?
// Or maybe it excludes test rows?
const nonTest = data2.filter(r => r['非測試'] === '非測試');
console.log(`\nFiltered 非測試: ${nonTest.length}`);

// Try filtering by '應回覆 == 1' and '非測試'
const finalFilter = data2.filter(r => r['非測試'] === '非測試' && r['應回覆'] == 1);
console.log(`Filtered 非測試 && 應回覆: ${finalFilter.length}`);

// Calculate rate and days for the nonTest set
const count1 = nonTest.length;
const solved1 = nonTest.filter(r => r['已回覆'] == 1).length;
const days1 = nonTest.reduce((s, r) => s + (Number(r['回覆天數']) || 0), 0) / count1;
console.log(`非測試 Rate: ${(solved1/count1)*100}%, Avg Days: ${days1}`);
