const path = require('path');
const { runAnalysisGenerator, fetchLiveApiData } = require('../scripts/generate-agent-analysis');

module.exports = async (req, res) => {
    // Enable CORS
    res.setHeader('Access-Control-Allow-Credentials', 'true');
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'GET,OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', 'X-CSRF-Token, X-Requested-With, Accept, Accept-Version, Content-Length, Content-MD5, Content-Type, Date, X-Api-Version');

    if (req.method === 'OPTIONS') {
        res.status(200).end();
        return;
    }

    try {
        const excelPath = path.join(__dirname, '..', 'MM2026_pistelaskenta.xlsx');
        const txtPath = path.join(__dirname, '..', 'tulokset.txt');

        console.log('Fetching live API data for dynamic serverless HTML request...');
        const { apiGames, apiGroups, apiTeams } = await fetchLiveApiData();

        const { finalJson } = await runAnalysisGenerator({
            excelPath,
            txtPath,
            apiGames,
            apiGroups,
            apiTeams,
            writeFiles: false
        });

        // Set caching headers to enable Vercel Edge caching
        res.setHeader('Cache-Control', 'public, max-age=0, s-maxage=300, stale-while-revalidate=60');
        res.setHeader('Content-Type', 'text/html; charset=utf-8');
        res.status(200).send(`<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <title>Veikkaus 2026 Agent Data</title>
</head>
<body>
    <pre id="agent-data" style="word-wrap: break-word; white-space: pre-wrap;">${JSON.stringify(finalJson)}</pre>
</body>
</html>`);
    } catch (err) {
        console.error('Serverless error generating analysis HTML:', err);
        res.setHeader('Content-Type', 'text/html; charset=utf-8');
        res.status(500).send(`<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <title>Error</title>
</head>
<body>
    <h1>Error</h1>
    <p>${err.message}</p>
</body>
</html>`);
    }
};
