const http = require('http');
const fs = require('fs');
const path = require('path');
const { runAnalysisGenerator, fetchLiveApiData } = require('./generate-agent-analysis');

const PORT = 3000;
const PUBLIC_DIR = path.join(__dirname, '..', 'public');

const server = http.createServer(async (req, res) => {
    // Enable CORS headers
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'GET,OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', '*');

    if (req.method === 'OPTIONS') {
        res.writeHead(200);
        res.end();
        return;
    }

    const url = new URL(req.url, `http://${req.headers.host}`);
    const pathname = url.pathname;

    // Handle agent-analysis.json dynamically (like Vercel)
    if (pathname === '/agent-analysis.json') {
        try {
            console.log('[DevServer] Dynamically generating agent-analysis.json...');
            const excelPath = path.join(__dirname, '..', 'MM2026_pistelaskenta.xlsx');
            const txtPath = path.join(__dirname, '..', 'tulokset.txt');
            
            const { apiGames, apiGroups, apiTeams } = await fetchLiveApiData();
            const { finalJson } = await runAnalysisGenerator({
                excelPath,
                txtPath,
                apiGames,
                apiGroups,
                apiTeams,
                writeFiles: false
            });

            res.writeHead(200, { 'Content-Type': 'application/json' });
            res.end(JSON.stringify(finalJson));
        } catch (err) {
            console.error('[DevServer] Error generating analysis:', err);
            res.writeHead(500, { 'Content-Type': 'application/json' });
            res.end(JSON.stringify({ error: err.message }));
        }
        return;
    }

    // Handle botmanen-summary.json dynamically (like Vercel)
    if (pathname === '/botmanen-summary.json') {
        try {
            console.log('[DevServer] Dynamically generating botmanen-summary.json...');
            const excelPath = path.join(__dirname, '..', 'MM2026_pistelaskenta.xlsx');
            const txtPath = path.join(__dirname, '..', 'tulokset.txt');
            
            const { apiGames, apiGroups, apiTeams } = await fetchLiveApiData();
            const { botmanenJson } = await runAnalysisGenerator({
                excelPath,
                txtPath,
                apiGames,
                apiGroups,
                apiTeams,
                writeFiles: false
            });

            res.writeHead(200, { 'Content-Type': 'application/json' });
            res.end(JSON.stringify(botmanenJson));
        } catch (err) {
            console.error('[DevServer] Error generating summary:', err);
            res.writeHead(500, { 'Content-Type': 'application/json' });
            res.end(JSON.stringify({ error: err.message }));
        }
        return;
    }

    // Handle live-games.json dynamically by proxying to the live API
    if (pathname === '/live-games.json') {
        try {
            console.log('[DevServer] Fetching live games from API...');
            const apiRes = await fetch('https://worldcup26.ir/get/games');
            if (!apiRes.ok) throw new Error(`API HTTP ${apiRes.status}`);
            const data = await apiRes.json();

            res.writeHead(200, { 'Content-Type': 'application/json', 'Cache-Control': 'no-store, no-cache, must-revalidate, max-age=0' });
            res.end(JSON.stringify(data));
        } catch (err) {
            console.error('[DevServer] Error fetching live games:', err);
            res.writeHead(500, { 'Content-Type': 'application/json' });
            res.end(JSON.stringify({ error: err.message }));
        }
        return;
    }

    // Serve static files from PUBLIC_DIR
    let filePath = path.join(PUBLIC_DIR, pathname === '/' ? 'index.html' : pathname);
    
    // Check if file exists
    if (!fs.existsSync(filePath) || fs.statSync(filePath).isDirectory()) {
        res.writeHead(404, { 'Content-Type': 'text/plain' });
        res.end('404 Not Found');
        return;
    }

    // Determine content type
    const ext = path.extname(filePath).toLowerCase();
    const mimeTypes = {
        '.html': 'text/html',
        '.css': 'text/css',
        '.js': 'text/javascript',
        '.json': 'application/json',
        '.png': 'image/png',
        '.jpg': 'image/jpeg',
        '.gif': 'image/gif',
        '.svg': 'image/svg+xml',
        '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        '.txt': 'text/plain'
    };
    const contentType = mimeTypes[ext] || 'application/octet-stream';

    // Read and serve file
    fs.readFile(filePath, (err, data) => {
        if (err) {
            res.writeHead(500, { 'Content-Type': 'text/plain' });
            res.end('500 Internal Server Error');
            return;
        }
        res.writeHead(200, { 
            'Content-Type': contentType,
            'Cache-Control': 'no-store, no-cache, must-revalidate, max-age=0' // Prevent caching locally
        });
        res.end(data);
    });
});

server.listen(PORT, () => {
    console.log(`\n==================================================`);
    console.log(`[DevServer] Local development server running at:`);
    console.log(`  http://localhost:${PORT}`);
    console.log(`[DevServer] Serving static files from: ${PUBLIC_DIR}`);
    console.log(`[DevServer] Dynamically routing /agent-analysis.json`);
    console.log(`==================================================\n`);
});
