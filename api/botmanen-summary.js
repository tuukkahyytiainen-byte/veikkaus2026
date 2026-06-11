const path = require('path');
const { runAnalysisGenerator, fetchLiveApiData } = require('../scripts/generate-agent-analysis');

module.exports = async (req, res) => {
    // Enable CORS
    res.setHeader('Access-Control-Allow-Credentials', 'true');
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'GET,OPTIONS,PATCH,DELETE,POST,PUT');
    res.setHeader('Access-Control-Allow-Headers', 'X-CSRF-Token, X-Requested-With, Accept, Accept-Version, Content-Length, Content-MD5, Content-Type, Date, X-Api-Version');

    if (req.method === 'OPTIONS') {
        res.status(200).end();
        return;
    }

    try {
        const excelPath = path.join(process.cwd(), 'MM2026_pistelaskenta.xlsx');
        const txtPath = path.join(process.cwd(), 'tulokset.txt');

        console.log('Fetching live API data for dynamic serverless request...');
        const { apiGames, apiGroups, apiTeams } = await fetchLiveApiData();

        const { botmanenJson } = await runAnalysisGenerator({
            excelPath,
            txtPath,
            apiGames,
            apiGroups,
            apiTeams,
            writeFiles: false
        });

        // Set caching headers: cache for 1 minute, allow stale-while-revalidate for 5 minutes
        res.setHeader('Cache-Control', 's-maxage=60, stale-while-revalidate=300');
        res.status(200).json(botmanenJson);
    } catch (err) {
        console.error('Serverless error generating Botmanen summary:', err);
        res.status(500).json({ error: err.message });
    }
};
