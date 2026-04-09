const XLSX = require('xlsx');
const workbook = XLSX.readFile('LINE OA 加入資料_分析結果.xlsx');

const sheet2 = workbook.Sheets['Data_回覆時程'];
const data2 = XLSX.utils.sheet_to_json(sheet2);

console.log("Sheet 2 Header Sample Data:");
console.log(data2.slice(0, 3).map(r => ({
    DEPT_NM: r.DEPT_NM, 
    AREA_NM: r.AREA_NM, 
    AGENCY_NAME: r.AGENCY_NAME, 
    AREA: r.AREA,
    應回覆: r.應回覆,
    已回覆: r.已回覆,
    三內: r['三天內回覆']
})));

// Check which column matches "AP6B", "F18B" etc.
const areaValues = [...new Set(data2.map(r => r.AREA))];
console.log("\nUnique values in AREA column (top 15):", areaValues.slice(0, 15));

const areaNmValues = [...new Set(data2.map(r => r.AREA_NM))];
console.log("Unique values in AREA_NM column (top 15):", areaNmValues.slice(0, 15));

// Find where 175 might come from
const counts = {
    total: data2.length,
    ni_ceshi: data2.filter(r => r['非測試'] === '非測試').length,
    ying_huifu: data2.filter(r => r['應回覆'] == 1).length,
    both: data2.filter(r => r['非測試'] === '非測試' && r['應回覆'] == 1).length
};
console.log("\nCounts:", counts);

// Check if 175 is the sum of rows where AGENCY_NAME is not empty and some other filter
const validAgencies = data2.filter(r => r.AGENCY_NAME && r['非測試'] === '非測試');
console.log("Valid Non-test agencies count:", validAgencies.length);
