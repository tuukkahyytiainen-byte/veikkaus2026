const XLSX = require('xlsx');
const path = require('path');

const excelFile = path.join(__dirname, 'MM2026_pistelaskenta.xlsx');

try {
    const wb = XLSX.readFile(excelFile);
    let foundComments = [];
    wb.SheetNames.forEach(name => {
        const sheet = wb.Sheets[name];
        for (const ref in sheet) {
            const cell = sheet[ref];
            if (cell && cell.c && cell.c.length > 0) {
                cell.c.forEach(comment => {
                    const text = comment.t || '';
                    foundComments.push({
                        sheet: name,
                        cell: ref,
                        text: text
                    });
                });
            }
        }
    });
    console.log(`Found ${foundComments.length} cell comments in the workbook.`);
    if (foundComments.length > 0) {
        console.log('Sample comments:', foundComments.slice(0, 10));
        fs.writeFileSync(path.join(__dirname, 'excel_cell_comments.txt'), JSON.stringify(foundComments, null, 2));
    }
} catch (err) {
    console.error('Error:', err);
}
