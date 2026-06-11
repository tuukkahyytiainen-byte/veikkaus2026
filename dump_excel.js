const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

const excelFile = path.join(__dirname, 'MM2026_pistelaskenta.xlsx');

try {
    const wb = XLSX.readFile(excelFile);
    let out = '';
    wb.SheetNames.forEach(name => {
        const sheet = wb.Sheets[name];
        out += `\n\n=== SHEET: ${name} ===\n`;
        for (const ref in sheet) {
            if (sheet[ref] && sheet[ref].v !== undefined) {
                out += `${ref}: ${sheet[ref].v}\n`;
            }
        }
    });
    fs.writeFileSync(path.join(__dirname, 'excel_dump.txt'), out);
    console.log('Successfully wrote dump to excel_dump.txt');
} catch (err) {
    console.error('Error:', err);
}
