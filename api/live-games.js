const https = require('https');

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

    https.get('https://worldcup26.ir/get/games', (apiRes) => {
        let data = '';

        apiRes.on('data', (chunk) => {
            data += chunk;
        });

        apiRes.on('end', () => {
            try {
                if (apiRes.statusCode !== 200) {
                    throw new Error(`API HTTP ${apiRes.statusCode}`);
                }
                const jsonData = JSON.parse(data);
                
                // Cache for 30 seconds at edge, allow stale-while-revalidate for 60 seconds
                res.setHeader('Cache-Control', 's-maxage=30, stale-while-revalidate=60');
                res.status(200).json(jsonData);
            } catch (err) {
                console.error('Error parsing live games JSON:', err);
                res.status(500).json({ error: err.message });
            }
        });
    }).on('error', (err) => {
        console.error('Error fetching live games serverless:', err);
        res.status(500).json({ error: err.message });
    });
};
