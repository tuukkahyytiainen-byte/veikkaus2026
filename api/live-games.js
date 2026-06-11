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
        console.log('Fetching live games from worldcup26.ir/get/games...');
        const apiRes = await fetch('https://worldcup26.ir/get/games');
        if (!apiRes.ok) throw new Error(`API HTTP ${apiRes.status}`);
        const data = await apiRes.json();
        
        // Cache for 30 seconds at edge, allow stale-while-revalidate for 60 seconds
        res.setHeader('Cache-Control', 's-maxage=30, stale-while-revalidate=60');
        res.status(200).json(data);
    } catch (err) {
        console.error('Error fetching live games serverless:', err);
        res.status(500).json({ error: err.message });
    }
};
