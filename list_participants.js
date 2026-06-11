const XLSX = require('xlsx');
const path = require('path');

const excelFile = path.join(__dirname, 'MM2026_pistelaskenta.xlsx');

try {
    const wb = XLSX.readFile(excelFile);
    const names = [];
    wb.SheetNames.forEach(name => {
        const sheet = wb.Sheets[name];
        if (sheet && sheet['H2']) {
            names.push({ sheet: name, name: sheet['H2'].v });
        }
    });
    console.log('Participant Names:', names);
} catch (err) {
    console.error('Error:', err);
}
