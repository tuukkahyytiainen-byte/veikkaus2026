const XLSX = require('xlsx');
const path = require('path');

const excelFile = path.join(__dirname, 'MM2026_pistelaskenta.xlsx');

try {
    const wb = XLSX.readFile(excelFile);
    wb.Workbook.Sheets.forEach((sheetMeta, idx) => {
        const name = wb.SheetNames[idx];
        const state = sheetMeta.Hidden; // 0 = visible, 1 = hidden, 2 = very hidden
        console.log(`Sheet name: "${name}", Hidden state: ${state === 0 ? 'Visible' : state === 1 ? 'Hidden' : state === 2 ? 'Very Hidden' : state}`);
        
        // Let's search inside this sheet for "Kaikki" or "Meksiko" or "kommentti"
        const sheet = wb.Sheets[name];
        let matchCount = 0;
        for (const ref in sheet) {
            if (sheet[ref] && sheet[ref].v !== undefined) {
                const val = String(sheet[ref].v);
                if (val.includes('1960') || val.includes('Meksiko – Etelä-Afrikka') || val.includes('lähtee Meksikon') || val.includes('läh')) {
                    console.log(`  [MATCH] cell ${ref}: "${val}"`);
                    matchCount++;
                }
            }
        }
    });
} catch (err) {
    console.error('Error:', err);
}
