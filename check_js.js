const fs = require('fs');
const vm = require('vm');

const html = fs.readFileSync('index.html', 'utf8');
const parts = html.split('<script>');
const jsPart = parts.find(p => p.includes('workbook = null'));
const jsCode = jsPart.substring(0, jsPart.indexOf('</script>'));
fs.writeFileSync('extracted_script.js', jsCode);

try {
    new vm.Script(jsCode);
    console.log("Syntax OK");
} catch(e) {
    console.error(e);
}
