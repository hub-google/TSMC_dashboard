const XLSX = require('xlsx');
const workbook = XLSX.readFile('LINE OA 加入資料_分析結果.xlsx');

const sheet1 = workbook.Sheets['各管道每日加入人數'];
const fullData1 = XLSX.utils.sheet_to_json(sheet1, { header: 1 });

console.log("Sheet 1 First 20 rows:");
console.log(fullData1.slice(0, 20));

console.log("\nSheet 1 Search for '2,074':");
fullData1.forEach((row, idx) => {
    if (row.includes(2074) || row.includes("2,074") || row.includes(2073)) {
        console.log(`Found near row ${idx}:`, row);
    }
});

const sheet2 = workbook.Sheets['Data_回覆時程'];
const data2 = XLSX.utils.sheet_to_json(sheet2);
console.log("\nSheet 2 Checking filters for 175:");
const filters = [
    { name: '非測試', value: '非測試' },
    { name: '應回覆', value: 1 },
    { name: '已回覆', value: 1 }
];

filters.forEach(f => {
    const matched = data2.filter(r => r[f.name] == f.value);
    console.log(`Filter ${f.name} == ${f.value}: ${matched.length}`);
});

const combined = data2.filter(r => r['非測試'] === '非測試' && r['應回覆'] == 1);
console.log(`Filter 非測試 && 應回覆 == 1: ${combined.length}`);
