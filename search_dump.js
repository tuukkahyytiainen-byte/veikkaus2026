const fs = require('fs');
const path = require('path');

const dumpFile = path.join(__dirname, 'excel_dump.txt');
const content = fs.readFileSync(dumpFile, 'utf8');
const lines = content.split('\n');

console.log('=== SEARCH RESULTS ===');
lines.forEach((line, idx) => {
    if (line.includes('1960') || line.includes('läh') || line.includes('Meksiko') || line.includes('Tshekki')) {
        console.log(`Line ${idx + 1}: ${line}`);
    }
});
