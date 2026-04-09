const XLSX = require('xlsx');
const workbook = XLSX.readFile('LINE OA 加入資料_分析結果.xlsx');

const sheet2 = workbook.Sheets['Data_回覆時程'];
const data2 = XLSX.utils.sheet_to_json(sheet2);

console.log("Searching for 175...");

// Try different filters
const filters = [
    { name: '非測試 only', f: r => r['非測試'] === '非測試' },
    { name: '應回覆 only', f: r => r['應回覆'] == 1 },
    { name: '非測試 & 應回覆', f: r => r['非測試'] === '非測試' && r['應回覆'] == 1 },
    { name: '已建檔 only', f: r => r['是否建檔'] },
    { name: '非測試 & 已建檔', f: r => r['非測試'] === '非測試' && r['是否建檔'] },
    { name: '非測試 & 應回覆 & 已建檔', f: r => r['非測試'] === '非測試' && r['應回覆'] == 1 && r['是否建檔'] }
];

filters.forEach(filter => {
    const subset = data2.filter(filter.f);
    const count = subset.length;
    const solved = subset.filter(r => r['已回覆'] == 1).length;
    const rate = (solved / count) * 100;
    const days = subset.reduce((s, r) => s + (Number(r['回覆天數']) || 0), 0) / count;
    
    console.log(`${filter.name}: Count=${count}, Rate=${rate.toFixed(1)}%, AvgDays=${days.toFixed(1)}`);
});

// Check AREA values for AP6B
const areaSet = new Set(data2.map(r => r.AREA));
console.log("\nIs 'AP6B' in AREA column?", areaSet.has('AP6B'));
console.log("Is '太陽' in AGENCY_NAME column?", new Set(data2.map(r => r.AGENCY_NAME)).has('太陽'));

// List a few rows where AREA starts with AP
console.log("\nRows where AREA starts with AP:");
console.log(data2.filter(r => r.AREA && r.AREA.startsWith('AP')).slice(0, 5).map(r => ({ AREA: r.AREA, AGENCY_NAME: r.AGENCY_NAME })));
