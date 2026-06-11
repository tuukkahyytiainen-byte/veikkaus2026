const XLSX = require('xlsx');
const path = require('path');

const excelFile = path.join(__dirname, 'MM2026_pistelaskenta.xlsx');

try {
    const wb = XLSX.readFile(excelFile);
    ['Pisteet', 'Laskenta'].forEach(name => {
        const sheet = wb.Sheets[name];
        console.log(`=== SHEET: ${name} ===`);
        let count = 0;
        for (const ref in sheet) {
            if (sheet[ref] && sheet[ref].v !== undefined) {
                console.log(`  ${ref}: ${sheet[ref].v}`);
                count++;
                if (count > 40) {
                    console.log('  ... and more');
                    break;
                }
            }
        }
    });
} catch (err) {
    console.error('Error:', err);
}
