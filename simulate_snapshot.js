const XLSX = require('xlsx');
const workbook = XLSX.readFile('LINE OA 加入資料_分析結果.xlsx');

const sheet1 = workbook.Sheets['各管道每日加入人數'];
const data1 = XLSX.utils.sheet_to_json(sheet1, { range: 1 });

// Filter up to March 11 (Excel serial number for 2026-03-11 is approx 46092)
// Wait, 46006 was late Dec? 45993 was Dec? 
// 46006 - 45993 = 13 days.
// 2026-04-09 is 46121? 
// Let's just find the row for 3/11.
const march11Rows = data1.filter(r => {
    if (typeof r['加入LINE OA日期'] !== 'number') return false;
    const dateStr = XLSX.utils.format_cell({ v: r['加入LINE OA日期'], t: 'd' });
    return dateStr.includes('3/11') || dateStr.includes('03/11');
});
console.log("March 11 found:", march11Rows.length > 0 ? march11Rows[0] : "Not found");

const cutoff = 46092; // Approx 3/11/2026.
const dataUpToMarch11 = data1.filter(r => typeof r['加入LINE OA日期'] === 'number' && r['加入LINE OA日期'] <= cutoff);

const hr = dataUpToMarch11.reduce((s, r) => s + (Number(r['由HR公告加入']) || 0), 0);
const friends = dataUpToMarch11.reduce((s, r) => s + (Number(r['由好友推薦加入']) || 0), 0);
const cards = dataUpToMarch11.reduce((s, r) => s + (Number(r['由服務小卡加入']) || 0), 0);
const totalInSheet = dataUpToMarch11.reduce((s, r) => s + (Number(r['總加入人數']) || 0), 0);

console.log(`\nSums up to approx 3/11:`);
console.log(`HR: ${hr}, Friend: ${friends}, Card: ${cards}, Total: ${totalInSheet}`);

// Let's check Sheet 2 with date filter
const sheet2 = workbook.Sheets['Data_回覆時程'];
const data2 = XLSX.utils.sheet_to_json(sheet2);
const data2UpToMarch11 = data2.filter(r => {
    const d = new Date(r['預約時間']);
    return d <= new Date('2026-03-11');
});
console.log(`\nSheet 2 rows up to 3/11: ${data2UpToMarch11.length}`);
console.log(`Sheet 2 '非測試' rows up to 3/11: ${data2UpToMarch11.filter(r => r['非測試'] === '非測試').length}`);
