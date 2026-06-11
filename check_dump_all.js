const fs = require('fs');
const path = require('path');

const file = path.join(__dirname, 'excel_dump.txt');
const raw = fs.readFileSync(file, 'utf8');

console.log('Search for "Kaikki":', raw.includes('Kaikki'));
console.log('Search for "Tips":', raw.includes('Tips'));
console.log('Search for "läh":', raw.includes('läh'));
console.log('Search for "kommentti":', raw.includes('kommentti'));

// Print lines containing "Kaikki"
const lines = raw.split('\n');
lines.forEach((line, idx) => {
    if (line.toLowerCase().includes('kaikki') || line.toLowerCase().includes('tips') || line.toLowerCase().includes('läh')) {
        console.log(`Line ${idx+1}: ${line}`);
    }
});
