const fs = require('fs');
const readline = require('readline');
const path = require('path');

const logFile = path.join('C:', 'Users', 'tuukk', '.gemini', 'antigravity', 'brain', '36f4bd21-e92f-4c98-a0b2-a5ce1a46fd87', '.system_generated', 'logs', 'transcript.jsonl');

async function search() {
    const fileStream = fs.createReadStream(logFile);
    const rl = readline.createInterface({
        input: fileStream,
        crlfDelay: Infinity
    });

    let index = 0;
    for await (const line of rl) {
        const parsed = JSON.parse(line);
        if (parsed.type === 'USER_INPUT') {
            console.log(`User Input at step ${parsed.step_index || index}:`);
            console.log(parsed.content.substring(0, 500));
            console.log('-----------------------------------');
        }
        index++;
    }
}

search();
