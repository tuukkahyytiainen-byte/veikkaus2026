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

    const options = {
        hostname: 'worldcup26.ir',
        path: '/get/games',
        method: 'GET',
        headers: {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
        }
    };

    https.get(options, (apiRes) => {
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
                
                // Cache for 10 seconds at edge, do not cache in browser, no stale-while-revalidate
                res.setHeader('Cache-Control', 'public, max-age=0, s-maxage=10');
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
