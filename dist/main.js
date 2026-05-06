// Configuration (Format-Only Alignment - FINAL SAFE VERSION)
const EXCEL_FILE_PATH = './LINE OA 加入資料_分析結果.xlsx';
const COLORS = {
    hr: '#0d325a',       // Deep Blue
    friends: '#d0af6b',  // Gold
    cards: '#929292',    // Grey
    sidebar: '#0056b3',
    progress: '#f97316'
};

let charts = {};

// Register Plugins
if (window.Chart) {
    if (window['chartjs-plugin-annotation']) Chart.register(window['chartjs-plugin-annotation']);
    if (window['ChartDataLabels']) Chart.register(window['ChartDataLabels']);
}

async function init() {
    await loadData();
    document.getElementById('refresh-btn').onclick = loadData;
}

const parseNum = (val) => {
    if (val === undefined || val === null || val === '') return 0;
    if (typeof val === 'number') return val;
    return parseFloat(val.toString().replace(/,/g, '')) || 0;
};

async function loadData() {
    try {
        const response = await fetch(EXCEL_FILE_PATH);
        const arrayBuffer = await response.arrayBuffer();
        const data = new Uint8Array(arrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });
        
        const updateTime = new Date();
        document.getElementById('update-ts').innerText = `${updateTime.toLocaleDateString()} ${updateTime.toLocaleTimeString()}`;

        const joinSheet = workbook.Sheets['各管道每日加入人數'];
        const respSheet = workbook.Sheets['Data_回覆時程'];

        if (joinSheet) processSlide1(joinSheet);
        if (respSheet) processSlide2(respSheet);
        if (joinSheet && respSheet) processSlide3(joinSheet, respSheet);
        
    } catch (error) {
        console.error('Data Load Error:', error);
    }
}

function processSlide1(sheet) {
    const allData = XLSX.utils.sheet_to_json(sheet, { range: 1 });
    const dailyData = allData.filter(row => row['加入LINE OA日期'] && typeof row['加入LINE OA日期'] === 'number');

    if (dailyData.length > 0) {
        const maxSerial = Math.max(...dailyData.map(r => r['加入LINE OA日期']));
        const maxDate = new Date((maxSerial - 25569) * 86400 * 1000);
        const dateStr = maxDate.getFullYear() + '/' + String(maxDate.getMonth()+1).padStart(2, '0') + '/' + String(maxDate.getDate()).padStart(2, '0');
        const statEl = document.getElementById('slide1-date-stat');
        if (statEl) statEl.innerText = '統計至: ' + dateStr;
    }

    const totalJoins = dailyData.reduce((s, r) => s + parseNum(r['總加入人數']), 0);
    document.getElementById('total-joins-large').innerText = totalJoins.toLocaleString();

    const hr = dailyData.reduce((s, r) => s + parseNum(r['由HR公告加入']), 0);
    const fri = dailyData.reduce((s, r) => s + parseNum(r['由好友推薦加入']), 0);
    const car = dailyData.reduce((s, r) => s + parseNum(r['由服務小卡加入']), 0);
    
    const getPctStr = (val) => ((val / totalJoins) * 100).toFixed(1) + '%';
    const names = ['HR 公告', '好友推薦', '服務小卡'];
    const counts = [hr, fri, car];

    renderChart('joinDonutChart', 'doughnut', names, [{
        data: counts,
        backgroundColor: [COLORS.hr, COLORS.friends, COLORS.cards],
        borderWidth: 0,
        datalabels: {
            color: '#1a202c',
            anchor: 'end',
            align: 'end',
            offset: 20,
            font: { size: 14, weight: '800' },
            formatter: (val, ctx) => {
                const name = ctx.chart.data.labels[ctx.dataIndex];
                return `${name}\n(${val.toLocaleString()}人)\n${getPctStr(val)}`;
            }
        }
    }], {
        plugins: { 
            legend: { display: false },
            datalabels: { display: true }
        },
        layout: { padding: { left: 80, right: 150, top: 20, bottom: 20 } },
        cutout: '55%'
    });

    const labels = dailyData.map(row => {
        const serial = row['加入LINE OA日期'];
        const date = new Date((serial - 25569) * 86400 * 1000);
        return `${date.getMonth() + 1}/${date.getDate()}`;
    });

    const datasets = [
        { label: 'HR 公告', data: dailyData.map(r => parseNum(r['由HR公告加入'])), backgroundColor: COLORS.hr },
        { label: '好友推薦', data: dailyData.map(r => parseNum(r['由好友推薦加入'])), backgroundColor: COLORS.friends },
        { label: '服務小卡', data: dailyData.map(r => parseNum(r['由服務小卡加入'])), backgroundColor: COLORS.cards }
    ].map(ds => ({ ...ds, fill: true, tension: 0.1, pointRadius: 0 }));

    const annotations = {};
    const peakDefs = [ 
        { d: '12/15', t: 'HR 公告發布首日', xOff: 180, yOff: 50 }, // MOVED DOWN AND RIGHT INTO SAFE ZONE
        { d: '12/18', t: '開始好友推薦', xOff: 180, yOff: 150 } 
    ];
    
    peakDefs.forEach((p, i) => {
        const idx = labels.indexOf(p.d);
        if (idx !== -1) {
            const val = dailyData[idx]['總加入人數'];
            annotations[`peakLabel${i}`] = {
                type: 'label',
                xValue: idx,
                yValue: val,
                content: [`高峰：${val}`, `(${p.d} ${p.t})`],
                backgroundColor: 'white',
                borderColor: '#d32f2f',
                borderWidth: 1.5,
                borderRadius: 4,
                padding: 10,
                font: { size: 12, weight: 'bold' },
                position: 'center',
                xAdjust: p.xOff, 
                yAdjust: p.yOff,
                callout: {
                    display: true,
                    borderColor: '#d32f2f',
                    borderWidth: 1.5,
                    side: 10
                },
                shadowBlur: 10,
                shadowColor: 'rgba(0,0,0,0.2)'
            };
            annotations[`peakPoint${i}`] = {
                type: 'point',
                xValue: idx,
                yValue: val,
                backgroundColor: 'white',
                borderColor: '#d32f2f',
                borderWidth: 2,
                radius: 6
            };
        }
    });

    renderChart('joinAreaChart', 'line', labels, datasets, {
        scales: {
            x: { 
                stacked: true, grid: { display: false },
                ticks: { autoSkip: true, maxTicksLimit: 14, font: { size: 11 } }
            },
            y: { 
                stacked: true, beginAtZero: true, 
                title: { display: true, text: '新加入用戶', font: { size: 13, weight: 'bold' }, padding: 10 },
                ticks: { font: { size: 11 } }
            }
        },
        plugins: {
            legend: { position: 'bottom', labels: { boxWidth: 12, padding: 20, font: { size: 12 } } },
            datalabels: { display: false },
            annotation: { 
                annotations,
                clip: false 
            }
        },
        layout: { padding: { right: 80, top: 20, bottom: 20 } }
    });
}

function processSlide2(sheet) {
    const data = XLSX.utils.sheet_to_json(sheet).filter(r => r['非測試'] === '非測試');
    
    // Global Header Stats
    const uniqueCustomers = new Set(data.filter(r => r.CUSTOMER_UUID).map(r => r.CUSTOMER_UUID)).size;
    const globalTargetRows = data.filter(r => parseNum(r['應回覆']) === 1);
    const globalSolvedRows = globalTargetRows.filter(r => parseNum(r['已回覆']) === 1);
    const overallRate = globalTargetRows.length > 0 ? (globalTargetRows.filter(r => parseNum(r['已回覆']) === 1).length / globalTargetRows.length) * 100 : 0;
    const globalAvgDays = globalSolvedRows.length > 0 ? globalSolvedRows.reduce((s, r) => s + parseNum(r['回覆天數']), 0) / globalSolvedRows.length : 0;

    document.getElementById('total-reservations').innerText = uniqueCustomers.toLocaleString();
    document.getElementById('overall-resp-rate').innerText = `${overallRate.toFixed(1)}%`;
    document.getElementById('avg-resp-days-sidebar').innerText = `${globalAvgDays.toFixed(1)}天`;

    const groups = {};
    data.forEach(row => {
        let areaKey = row.AREA || 'Unknown';
        if (areaKey.includes('其他廠區')) areaKey = '其他廠區';

        if (!groups[areaKey]) {
            groups[areaKey] = { 
                area: areaKey, 
                agencies: new Set(), 
                uuids: new Set(), 
                daysSum: 0, 
                solvedRows: 0, 
                targetRows: 0, 
                count3d: 0 
            };
        }
        
        const g = groups[areaKey];
        if (row.AGENCY_NAME) g.agencies.add(row.AGENCY_NAME);
        if (row.CUSTOMER_UUID) g.uuids.add(row.CUSTOMER_UUID);
        
        if (parseNum(row['應回覆']) === 1) {
            g.targetRows++;
            if (parseNum(row['三天內回覆']) === 1) g.count3d++;
            if (parseNum(row['已回覆']) === 1) {
                g.solvedRows++;
                g.daysSum += parseNum(row['回覆天數']);
            }
        }
    });

    const tableBody = document.getElementById('table-body');
    tableBody.innerHTML = '';

    Object.values(groups).sort((a,b) => {
        if (a.area === '其他廠區') return 1;
        if (b.area === '其他廠區') return -1;
        return b.uuids.size - a.uuids.size;
    }).forEach(item => {
        const rate = item.targetRows > 0 ? Math.round((item.count3d / item.targetRows) * 100) : 0;
        const avgDaysText = item.solvedRows > 0 ? (item.daysSum / item.solvedRows).toFixed(1) : '-';
        const agencyText = item.agencies.size > 0 ? Array.from(item.agencies).join(' / ') : '-';

        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td>${item.area}</td>
            <td>${agencyText}</td>
            <td>${item.uuids.size}</td>
            <td>${avgDaysText}</td>
            <td>
                <div class="progress-cell">
                    <div class="progress-fill" style="width: ${rate}%"></div>
                    <div class="progress-text">${rate}%</div>
                </div>
            </td>
        `;
        tableBody.appendChild(tr);
    });
}

function processSlide3(joinSheet, respSheet) {
    const joinData = XLSX.utils.sheet_to_json(joinSheet, { range: 1 });
    const respData = XLSX.utils.sheet_to_json(respSheet).filter(r => r['非測試'] === '非測試');

    const getMonthStr = (val) => {
        let d;
        if (typeof val === 'number') d = new Date((val - 25569) * 86400 * 1000);
        else d = new Date(val);
        if (isNaN(d.getTime())) return null;
        return `${d.getFullYear()}/${(d.getMonth() + 1).toString().padStart(2, '0')}`;
    };

    // 1. 加入好友數 (依月份)
    const joinByMonth = {};
    joinData.forEach(r => {
        const m = getMonthStr(r['加入LINE OA日期']);
        if (m) {
            joinByMonth[m] = (joinByMonth[m] || 0) + parseNum(r['總加入人數']);
        }
    });
    const joinMonths = Object.keys(joinByMonth).sort();
    renderChart('friendMonthChart', 'bar', joinMonths, [{
        label: '每月加入人數',
        data: joinMonths.map(m => joinByMonth[m]),
        backgroundColor: '#3182ce', // Lighter corporate blue for better contrast
        borderRadius: 5
    }], { scales: { y: { beginAtZero: true } } });

    // 2. 預約顧問人數 (依月份)
    const bookingByMonth = {};
    const bookingUniqueMonths = {}; // To store unique UUIDs per month
    respData.forEach(r => {
        const m = getMonthStr(r['預約時間']);
        if (m && r.CUSTOMER_UUID) {
            if (!bookingUniqueMonths[m]) bookingUniqueMonths[m] = new Set();
            bookingUniqueMonths[m].add(r.CUSTOMER_UUID);
        }
    });
    const bookingMonths = Object.keys(bookingUniqueMonths).sort();
    renderChart('bookingMonthChart', 'bar', bookingMonths, [{
        label: '每月預約人數',
        data: bookingMonths.map(m => bookingUniqueMonths[m].size),
        backgroundColor: COLORS.friends,
        borderRadius: 5
    }], { scales: { y: { beginAtZero: true } } });

    // 3. 駐廠回覆率 (依月份)
    const respRateByMonth = {};
    respData.filter(r => parseNum(r['應回覆']) === 1).forEach(r => {
        const m = getMonthStr(r['預約時間']);
        if (m) {
            if (!respRateByMonth[m]) respRateByMonth[m] = { total: 0, ok: 0 };
            respRateByMonth[m].total++;
            // 只要有回覆就算 (使用已回覆欄位)
            if (parseNum(r['已回覆']) === 1) respRateByMonth[m].ok++;
        }
    });
    const rateMonths = Object.keys(respRateByMonth).sort();
    renderChart('responseRateMonthChart', 'line', rateMonths, [{
        label: '駐廠回覆率 (%)',
        data: rateMonths.map(m => Math.round((respRateByMonth[m].ok / respRateByMonth[m].total) * 100)),
        borderColor: COLORS.progress,
        backgroundColor: 'rgba(249, 115, 22, 0.1)',
        fill: true,
        tension: 0.3,
        pointRadius: 5,
        pointBackgroundColor: COLORS.progress
    }], { 
        scales: { 
            y: { 
                beginAtZero: false, 
                max: 100,
                ticks: {
                    callback: (val) => val + '%'
                }
            } 
        } 
    });
}

function renderChart(id, type, labels, datasets, extraOptions = {}) {
    if (charts[id]) charts[id].destroy();
    const ctx = document.getElementById(id).getContext('2d');
    charts[id] = new Chart(ctx, {
        type, data: { labels, datasets },
        options: { responsive: true, maintainAspectRatio: false, ...extraOptions }
    });
}

init();
