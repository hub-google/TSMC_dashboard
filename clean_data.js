const XLSX = require('xlsx');

try {
    console.log('Reading data.xlsx...');
    const wb = XLSX.readFile('public/data.xlsx');
    
    // List of PII columns to remove
    const piiCols = ['CUST_NAME', 'LINE_UUID', 'SURNAME', 'AGENT_NUMBER'];
    
    wb.SheetNames.forEach(sheetName => {
        console.log(`Processing sheet: ${sheetName}`);
        const sheet = wb.Sheets[sheetName];
        
        // Convert to JSON, keeping empty cells so we don't lose structure
        const data = XLSX.utils.sheet_to_json(sheet, { defval: "" });
        
        let cleaned = false;
        data.forEach(row => {
            piiCols.forEach(col => {
                if (row[col] !== undefined) {
                    delete row[col];
                    cleaned = true;
                }
            });
        });
        
        if (cleaned) {
            console.log(`Cleaned PII from sheet: ${sheetName}`);
            // Convert back to sheet and replace the old one
            const newSheet = XLSX.utils.json_to_sheet(data);
            wb.Sheets[sheetName] = newSheet;
        }
    });
    
    console.log('Writing clean data back to data.xlsx...');
    XLSX.writeFile(wb, 'public/data.xlsx');
    console.log('Successfully removed PII from Excel file!');
} catch (err) {
    console.error('Error cleaning data:', err);
    process.exit(1);
}
