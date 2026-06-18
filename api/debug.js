const fs = require('fs');
const path = require('path');

module.exports = async (req, res) => {
    const debugInfo = {
        cwd: process.cwd(),
        __dirname: __dirname,
        env: {
            VERCEL: process.env.VERCEL,
            GITHUB_ACTIONS: process.env.GITHUB_ACTIONS
        },
        files: {}
    };

    const checkPaths = [
        path.join(process.cwd(), 'MM2026_pistelaskenta.xlsx'),
        path.join(process.cwd(), 'scripts', 'cached-api-data.json'),
        path.join(__dirname, 'MM2026_pistelaskenta.xlsx'),
        path.join(__dirname, '..', 'MM2026_pistelaskenta.xlsx'),
        path.join(__dirname, 'cached-api-data.json'),
        path.join(__dirname, '..', 'scripts', 'cached-api-data.json'),
        '/var/task/MM2026_pistelaskenta.xlsx',
        '/var/task/scripts/cached-api-data.json'
    ];

    checkPaths.forEach(p => {
        debugInfo.files[p] = {
            exists: fs.existsSync(p),
            isDir: fs.existsSync(p) ? fs.statSync(p).isDirectory() : false
        };
    });

    // List directory contents of process.cwd() and /var/task
    try {
        debugInfo.cwdContents = fs.readdirSync(process.cwd());
    } catch (e) {
        debugInfo.cwdContentsError = e.message;
    }

    try {
        debugInfo.taskContents = fs.readdirSync('/var/task');
    } catch (e) {
        debugInfo.taskContentsError = e.message;
    }

    try {
        debugInfo.apiContents = fs.readdirSync(path.join('/var/task', 'api'));
    } catch (e) {
        debugInfo.apiContentsError = e.message;
    }

    res.status(200).json(debugInfo);
};
