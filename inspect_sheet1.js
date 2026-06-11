const XLSX = require('xlsx');
const path = require('path');

const excelFile = path.join(__dirname, 'MM2026_pistelaskenta.xlsx');

try {
    const wb = XLSX.readFile(excelFile);
    const sheet = wb.Sheets['1'];
    // Let's print cells in rows 6 to 15
    for (let r = 5; r <= 15; r++) {
        let rowStr = `Row ${r}: `;
        ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P'].forEach(col => {
            const cell = sheet[`${col}${r}`];
            if (cell && cell.v !== undefined) {
                rowStr += `${col}:${cell.v} | `;
            }
        });
        console.log(rowStr);
    }
} catch (err) {
    console.error('Error:', err);
}
