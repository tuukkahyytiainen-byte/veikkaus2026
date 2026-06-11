const XLSX = require('xlsx');
const path = require('path');

const excelFile = path.join(__dirname, 'MM2026_pistelaskenta.xlsx');

try {
    const wb = XLSX.readFile(excelFile);
    let found = false;
    wb.SheetNames.forEach(name => {
        const sheet = wb.Sheets[name];
        for (const ref in sheet) {
            if (sheet[ref] && sheet[ref].v !== undefined) {
                const val = String(sheet[ref].v);
                if (val.startsWith('1960Tips') || val.startsWith('Kaikki läh') || val.startsWith('Kanada ko') || val.startsWith('USA:lle ko')) {
                    console.log(`FOUND in sheet "${name}" cell ${ref}: "${val}"`);
                    found = true;
                }
            }
        }
    });
    if (!found) {
        console.log('No cells match these comment prefixes in the workbook.');
    }
} catch (err) {
    console.error('Error:', err);
}
