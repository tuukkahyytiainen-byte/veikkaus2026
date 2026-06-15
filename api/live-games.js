const { fetchLiveApiData } = require('../scripts/generate-agent-analysis');

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
        const { apiGames } = await fetchLiveApiData();
        
        // Cache for 10 seconds at edge, do not cache in browser
        res.setHeader('Cache-Control', 'public, max-age=0, s-maxage=10');
        res.status(200).json({ games: apiGames });
    } catch (err) {
        console.error('Error fetching live games serverless:', err);
        res.status(500).json({ error: err.message });
    }
};
