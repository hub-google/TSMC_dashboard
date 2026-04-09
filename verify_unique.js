const XLSX = require('xlsx');
const workbook = XLSX.readFile('LINE OA 加入資料_分析結果.xlsx');

const sheet2 = workbook.Sheets['Data_回覆時程'];
const data2 = XLSX.utils.sheet_to_json(sheet2);

const nonTest = data2.filter(r => r['非測試'] === '非測試');
const uniqueCustomers = new Set(nonTest.map(r => r.CUSTOMER_UUID));

console.log("Filtered '非測試' Rows:", nonTest.length);
console.log("Unique 'CUSTOMER_UUID' Count:", uniqueCustomers.size);

// check if 175 matches unique count on 3/11
const data2UpToMarch11 = data2.filter(r => {
    const d = new Date(r['預約時間']);
    return d <= new Date('2026-03-11') && r['非測試'] === '非測試';
});
const uniqueCustomersM11 = new Set(data2UpToMarch11.map(r => r.CUSTOMER_UUID));
console.log("Unique 'CUSTOMER_UUID' Count up to March 11:", uniqueCustomersM11.size);
