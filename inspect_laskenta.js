const XLSX = require('xlsx');
const path = require('path');

const excelFile = path.join(__dirname, 'MM2026_pistelaskenta.xlsx');

try {
    const wb = XLSX.readFile(excelFile);
    ['Tulokset', 'Pisteet', 'Laskenta'].forEach(name => {
        const sheet = wb.Sheets[name];
        console.log(`=== Sheet: ${name} ===`);
        for (const ref in sheet) {
            if (sheet[ref] && sheet[ref].v !== undefined) {
                const val = String(sheet[ref].v);
                if (val.includes('Tips') || val.includes('Kaikki') || val.includes('kommentti') || val.includes('perustelu')) {
                    console.log(`  Cell ${ref}: "${val}"`);
                }
            }
        }
    });
} catch (err) {
    console.error('Error:', err);
}
