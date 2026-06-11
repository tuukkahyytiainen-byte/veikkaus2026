const XLSX = require('xlsx');
const path = require('path');

const excelFile = path.join(__dirname, 'MM2026_pistelaskenta.xlsx');

try {
    const wb = XLSX.readFile(excelFile);
    console.log('SheetNames:', wb.SheetNames);
    
    // Check if any sheet contains "Kaikki" or "Botmanen" or "1960Tips"
    let found = [];
    wb.SheetNames.forEach(name => {
        const sheet = wb.Sheets[name];
        let foundInSheet = false;
        for (const ref in sheet) {
            if (sheet[ref] && sheet[ref].v) {
                const val = String(sheet[ref].v);
                if (val.includes('1960Tips') || val.includes('lähtee Meksikon') || val.includes('Botmanen')) {
                    found.push({ sheet: name, cell: ref, val });
                }
            }
        }
    });
    console.log('Search matches:', found);
} catch (err) {
    console.error('Error:', err);
}
