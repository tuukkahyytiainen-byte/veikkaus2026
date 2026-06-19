const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');

// Normalization mappings copied from index.html
const apiTeamNameToFinnish = {
    "south africa": "Etelä-Afrikka",
    "brazil": "Brasilia",
    "scotland": "Skotlanti",
    "turkey": "Turkki",
    "ivory coast": "Norsunluurannikko",
    "netherlands": "Hollanti",
    "cape verde": "Kap Verde",
    "cape verde islands": "Kap Verde",
    "france": "Ranska",
    "tunisia": "Tunisia",
    "egypt": "Egypti",
    "iraq": "Irak",
    "portugal": "Portugali",
    "uzbekistan": "Uzbekistan",
    "colombia": "Kolumbia",
    "ecuador": "Ecuador",
    "japan": "Japani",
    "new zealand": "Uusi-Seelanti",
    "saudi arabia": "Saudi-Arabia",
    "austria": "Itävalta",
    "ghana": "Ghana",
    "south korea": "Etelä-Korea",
    "spain": "Espanja",
    "norway": "Norja",
    "argentina": "Argentiina",
    "democratic republic of the congo": "Kongon DR",
    "congo dr": "Kongon DR",
    "england": "Englanti",
    "czech republic": "Tshekki",
    "czechia": "Tshekki",
    "canada": "Kanada",
    "qatar": "Qatar",
    "switzerland": "Sveitsi",
    "morocco": "Marokko",
    "paraguay": "Paraguay",
    "curaçao": "Curaçao",
    "sweden": "Ruotsi",
    "algeria": "Algeria",
    "jordan": "Jordania",
    "haiti": "Haiti",
    "germany": "Saksa",
    "uruguay": "Uruguay",
    "senegal": "Senegal",
    "panama": "Panama",
    "mexico": "Meksiko",
    "bosnia and herzegovina": "Bosnia ja Hertsegovina",
    "bosnia-herzegovina": "Bosnia ja Hertsegovina",
    "bosnia herzegovina": "Bosnia ja Hertsegovina",
    "united states": "USA",
    "australia": "Australia",
    "belgium": "Belgia",
    "iran": "Iran",
    "croatia": "Kroatia"
};

function parseMatchDate(dateStr) {
    if (!dateStr) return 0;
    
    // Check if it is in MM/DD/YYYY HH:MM format
    if (dateStr.includes('/')) {
        const d = new Date(dateStr);
        if (!isNaN(d.getTime())) return d.getTime();
    }
    
    // Check if it is in D.M. klo HH format
    if (dateStr.includes(' klo ')) {
        const parts = dateStr.split(' klo ');
        const dateParts = parts[0].split('.');
        if (dateParts.length >= 2) {
            const day = parseInt(dateParts[0], 10);
            const month = parseInt(dateParts[1], 10);
            const hour = parseInt(parts[1], 10);
            return new Date(2026, month - 1, day, hour).getTime();
        }
    }
    
    // Fallback standard Date parsing
    const d = new Date(dateStr);
    return isNaN(d.getTime()) ? 0 : d.getTime();
}

function getLogicalDatePrefix(dateStr) {
    if (!dateStr) return "";
    const parts = dateStr.split(' klo ');
    if (parts.length < 2) return dateStr.split(' ')[0] || "";
    const dateParts = parts[0].split('.');
    if (dateParts.length < 2) return dateStr.split(' ')[0] || "";
    let day = parseInt(dateParts[0], 10);
    let month = parseInt(dateParts[1], 10);
    const hour = parseInt(parts[1], 10);
    
    // Shift early morning hours (0-6) to the previous day
    if (hour < 7) {
        const d = new Date(2026, month - 1, day);
        d.setDate(d.getDate() - 1);
        day = d.getDate();
        month = d.getMonth() + 1;
    }
    return `${day}.${month}.`;
}

function normalizeTeamName(name) {
    if (!name) return "";
    let clean = name.trim().toLowerCase()
        .replace(/[\.-]/g, ' ')
        .replace(/\s+/g, ' ');

    if (apiTeamNameToFinnish[clean]) return apiTeamNameToFinnish[clean];

    const mapping = {
        "algeria": "Algeria",
        "argentiina": "Argentiina", "argentina": "Argentiina", "argenttiina": "Argentiina",
        "australia": "Australia",
        "belgia": "Belgia",
        "bosnia hertegovina": "Bosnia ja Hertsegovina", 
        "bosnia hertse": "Bosnia ja Hertsegovina", 
        "bosnia hertzegovina": "Bosnia ja Hertsegovina",
        "bosnia ja hertsegovina": "Bosnia ja Hertsegovina", 
        "bosnia-hertsegovina": "Bosnia ja Hertsegovina",
        "bosnia herzegovina": "Bosnia ja Hertsegovina", 
        "bosnia": "Bosnia ja Hertsegovina",
        "brasialia": "Brasilia", "brasilia": "Brasilia",
        "curacao": "Curaçao", "curaçao": "Curaçao",
        "ecuador": "Ecuador", "equador": "Ecuador",
        "egypti": "Egypti", "eqypti": "Egypti",
        "englanti": "Englanti",
        "espanja": "Espanja",
        "etekä korea": "Etelä-Korea", "etela afrikka": "Etelä-Afrikka", "etelä afrikka": "Etelä-Afrikka",
        "etelä korea": "Etelä-Korea", "etelä-afrikka": "Etelä-Afrikka", "etelä-korea": "Etelä-Korea",
        "ghana": "Ghana",
        "haiti": "Haiti", "vaiti": "Haiti",
        "hollanti": "Hollanti",
        "irak": "Irak",
        "iran": "Iran",
        "itävalta": "Itävalta",
        "japani": "Japani",
        "jordania": "Jordania",
        "kanada": "Kanada",
        "kap verde": "Kap Verde",
        "cape verde": "Kap Verde",
        "cape verde islands": "Kap Verde",
        "kolumbia": "Kolumbia",
        "kongo": "Kongon DR", "kongon dr": "Kongon DR",
        "kroatia": "Kroatia",
        "marokko": "Marokko",
        "meksiko": "Meksiko", "mexico": "Meksiko",
        "norja": "Norja",
        "norsunluurannikko": "Norsunluurannikko", "norsunluurannikon": "Norsunluurannikko",
        "panama": "Panama",
        "paraguay": "Paraguay", "paraguaysta": "Paraguay", "paraquai": "Paraguay", "paraquay": "Paraguay",
        "portugal": "Portugali", "portugali": "Portugali",
        "qagar": "Qatar", "qatar": "Qatar", "quatar": "Qatar",
        "ranska": "Ranska",
        "ruotsi": "Ruotsi",
        "saksa": "Saksa",
        "saudi arabia": "Saudi-Arabia", "saudi-arabia": "Saudi-Arabia",
        "senegal": "Senegal", "senegas": "Senegal",
        "skotlanti": "Skotlanti",
        "sveitsi": "Sveitsi",
        "czechia": "Tshekki",
        "czech republic": "Tshekki",
        "tsekki": "Tshekki", "tshekki": "Tshekki", "tšekki": "Tshekki",
        "tunisia": "Tunisia",
        "turkki": "Turkki",
        "usa": "USA",
        "uruguay": "Uruguay", "uruqay": "Uruguay", "uruquai": "Uruguay",
        "uusi seelantia": "Uusi-Seelanti", "uusi seelanti": "Uusi-Seelanti", "uusi-seelanti": "Uusi-Seelanti",
        "uzbekistan": "Uzbekistan"
    };

    if (mapping[clean]) return mapping[clean];
    for (const key in mapping) {
        if (clean.includes(key) || key.includes(clean)) return mapping[key];
    }
    return name.charAt(0).toUpperCase() + name.slice(1);
}

function parseResultsText(text) {
    const lines = text.split('\n');
    const parsedMatches = {};
    const parsedStandings = {};
    let parsedFinalists = [];
    let parsed34 = [];

    lines.forEach(line => {
        line = line.trim();
        if (!line || line.startsWith('#')) return;
        const matchMatch = line.match(/^(\d+)\s*:\s*(\d+)\s*-\s*(\d+)/);
        if (matchMatch) {
            parsedMatches[Number(matchMatch[1])] = { goals1: Number(matchMatch[2]), goals2: Number(matchMatch[3]) };
            return;
        }
        const groupMatch = line.match(/^Lohko\s+([A-L])\s*:\s*(.+)$/i);
        if (groupMatch) {
            const teams = groupMatch[2].split(',').map(t => t.trim()).filter(Boolean);
            if (teams.length > 0) parsedStandings[groupMatch[1].toUpperCase()] = teams.map(t => normalizeTeamName(t));
            return;
        }
        const finalistsMatch = line.match(/^Finaali\s*:\s*(.+)$/i);
        if (finalistsMatch) parsedFinalists = finalistsMatch[1].split(',').map(t => normalizeTeamName(t.trim())).filter(Boolean);
        const thirdFourthMatch = line.match(/^Pronssi\s*:\s*(.+)$/i);
        if (thirdFourthMatch) parsed34 = thirdFourthMatch[1].split(',').map(t => normalizeTeamName(t.trim())).filter(Boolean);
    });
    return { matches: parsedMatches, standings: parsedStandings, finalists: parsedFinalists, thirdFourth: parsed34 };
}

function getGroupTeamsFromMatches(groupLetter, matchesList) {
    const teams = new Set();
    matchesList.forEach(m => {
        if (m.group === groupLetter) {
            if (m.team1) teams.add(normalizeTeamName(m.team1));
            if (m.team2) teams.add(normalizeTeamName(m.team2));
        }
    });
    return Array.from(teams);
}

function calculateGroupStandings(groupLetter, originalOrder, matchesList) {
    const groupMatches = matchesList.filter(m => m.group === groupLetter);
    const playedGroupMatches = groupMatches.filter(m => m.hasResult);
    
    if (playedGroupMatches.length === 0) return originalOrder;
    
    const stats = {};
    originalOrder.forEach(t => {
        if (t) {
            stats[t] = { name: t, points: 0, gd: 0, gf: 0, ga: 0, wins: 0, draws: 0, losses: 0 };
        }
    });
    
    groupMatches.forEach(m => {
        if (!m.hasResult) return;
        const t1 = m.team1;
        const t2 = m.team2;
        const g1 = m.goals1;
        const g2 = m.goals2;
        
        if (stats[t1] && stats[t2]) {
            stats[t1].gf += g1;
            stats[t1].ga += g2;
            stats[t1].gd += (g1 - g2);
            
            stats[t2].gf += g2;
            stats[t2].ga += g1;
            stats[t2].gd += (g2 - g1);
            
            if (g1 > g2) {
                stats[t1].points += 3;
                stats[t1].wins += 1;
                stats[t2].losses += 1;
            } else if (g1 < g2) {
                stats[t2].points += 3;
                stats[t2].wins += 1;
                stats[t1].losses += 1;
            } else {
                stats[t1].points += 1;
                stats[t1].draws += 1;
                stats[t2].points += 1;
                stats[t2].draws += 1;
            }
        }
    });
    
    const sortedTeams = Object.values(stats).sort((a, b) => {
        return b.points - a.points || 
               b.gd - a.gd || 
               b.gf - a.gf || 
               (originalOrder.indexOf(a.name) - originalOrder.indexOf(b.name));
    });
    return sortedTeams.map(t => t.name);
}

function calculateLeaderboard(customMatches, customStandings, customFinalists, custom34, workbook) {
    const pisteetSheet = workbook.Sheets['Pisteet'];
    const pisteetRows = XLSX.utils.sheet_to_json(pisteetSheet, { header: 1 });
    const calculatedPlayers = [];
    const groups = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L'];

    for (let r = 5; r < pisteetRows.length; r++) {
        const row = pisteetRows[r];
        if (!row || !row[1]) continue;
        const sheetName = String(row[1]).trim();
        const pSheet = workbook.Sheets[sheetName];
        if (!pSheet) continue;
        const getPValue = (addr) => pSheet[addr] ? pSheet[addr].v : undefined;
        const nameInSheet = getPValue('H2');
        const nameInRow = row[2];
        if ((!nameInSheet || !String(nameInSheet).trim()) && (!nameInRow || !String(nameInRow).trim())) {
            continue;
        }
        const participantName = getPValue('H2') ? String(getPValue('H2')).trim() : (row[2] ? String(row[2]).trim() : `Osallistuja ${sheetName}`);
        
        let matchPoints = 0, exactlyCorrect = 0;
        const playerMatches = [];

        for (let m = 1; m <= 72; m++) {
            const g1Val = getPValue(`I${6 + m}`), g2Val = getPValue(`K${6 + m}`);
            const actual = customMatches[m - 1], hasGuess = g1Val !== undefined && g1Val !== null && g2Val !== undefined && g2Val !== null;
            let pts = null;
            if (actual.hasResult && hasGuess) {
                const guess1 = Number(g1Val), guess2 = Number(g2Val), act1 = actual.goals1, act2 = actual.goals2;
                if (Math.sign(guess1 - guess2) !== Math.sign(act1 - act2)) pts = 0;
                else if (act1 === act2) pts = (guess1 === act1 && guess2 === act2) ? 6 : 3;
                else pts = Math.max(0, 6 - (Math.abs(guess1 - act1) + Math.abs(guess2 - act2)));
                matchPoints += pts;
                if (pts === 6) exactlyCorrect++;
            }
            playerMatches.push({ nr: m, guess1: hasGuess ? Number(g1Val) : null, guess2: hasGuess ? Number(g2Val) : null, points: pts, hasGuess: hasGuess });
        }

        let groupPoints = 0;
        const groupGuesses = {};
        groups.forEach((g, gIdx) => {
            groupGuesses[g] = [];
            const startRow = gIdx >= 6 ? 14 : 8;
            const colLetter = XLSX.utils.encode_col(15 + (gIdx % 6));
            for (let pos = 1; pos <= 4; pos++) {
                const guess = normalizeTeamName(String(getPValue(`${colLetter}${startRow + pos - 1}`) || ''));
                const actual = customStandings[g][pos - 1];
                let pts = (actual && guess === actual) ? 2 : 0;
                groupPoints += pts;
                groupGuesses[g].push({ pos, guess, actual, points: pts });
            }
        });

        const f1 = normalizeTeamName(String(getPValue('E79') || '')), f2 = normalizeTeamName(String(getPValue('G79') || ''));
        const t1 = normalizeTeamName(String(getPValue('E80') || '')), t2 = normalizeTeamName(String(getPValue('G80') || ''));
        
        let finalistPoints = 0;
        if (customFinalists.length > 0) {
            if (f1 && customFinalists.includes(f1)) finalistPoints += 2;
            if (f2 && customFinalists.includes(f2)) finalistPoints += 2;
        }
        let thirdFourthPoints = 0;
        if (custom34.length > 0) {
            if (t1 && custom34.includes(t1)) thirdFourthPoints += 2;
            if (t2 && custom34.includes(t2)) thirdFourthPoints += 2;
        }

        const allGroupsPlayed = customMatches.every(m => m.hasResult);
        const finalistsAdded = customFinalists.length === 2;
        const thirdFourthAdded = custom34.length === 2;

        const addedGroupPoints = allGroupsPlayed ? groupPoints : 0;
        const addedFinalistPoints = finalistsAdded ? finalistPoints : 0;
        const addedThirdFourthPoints = thirdFourthAdded ? thirdFourthPoints : 0;

        const totalPoints = matchPoints + addedGroupPoints + addedFinalistPoints + addedThirdFourthPoints;

        calculatedPlayers.push({
            sheetName,
            name: participantName,
            matchPoints,
            groupPoints,
            finalPoints: finalistPoints + thirdFourthPoints,
            totalPoints: totalPoints,
            exactlyCorrect,
            matches: playerMatches,
            groupGuesses,
            finalistGuesses: [f1, f2],
            thirdFourthGuesses: [t1, t2],
            groupPointsAdded: allGroupsPlayed,
            finalistsAdded: finalistsAdded,
            thirdFourthAdded: thirdFourthAdded
        });
    }

    calculatedPlayers.sort((a, b) => b.totalPoints - a.totalPoints || b.exactlyCorrect - a.exactlyCorrect || a.name.localeCompare(b.name));
    calculatedPlayers.forEach((p, i) => {
        p.rank = (i > 0 && p.totalPoints === calculatedPlayers[i - 1].totalPoints) ? calculatedPlayers[i - 1].rank : i + 1;
    });

    return calculatedPlayers;
}

function getLeaderboardAtState(upToMatchNr, matches, textResults, workbook) {
    const subMatches = matches.map(m => {
        const copy = { ...m };
        if (m.nr > upToMatchNr) {
            copy.goals1 = null;
            copy.goals2 = null;
            copy.hasResult = false;
            copy.status = 'Pelaamaton';
        }
        return copy;
    });

    const subStandings = {};
    const groups = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L'];
    const tuloksetSheet = workbook.Sheets['Tulokset'];
    const getTValue = (addr) => tuloksetSheet[addr] ? tuloksetSheet[addr].v : undefined;

    groups.forEach((g, gIdx) => {
        if (upToMatchNr >= 72 && textResults && textResults.standings && textResults.standings[g]) {
            subStandings[g] = textResults.standings[g];
        } else {
            const isSecondHalf = gIdx >= 6;
            const startRow = isSecondHalf ? 12 : 6;
            const colLetter = XLSX.utils.encode_col(12 + (gIdx % 6));
            const originalOrder = [];
            for (let pos = 1; pos <= 4; pos++) {
                const val = getTValue(`${colLetter}${startRow + pos - 1}`);
                originalOrder.push(val ? normalizeTeamName(String(val)) : null);
            }
            const cleanOrder = originalOrder.filter(Boolean);
            const fallbackOrder = cleanOrder.length === 4 ? cleanOrder : getGroupTeamsFromMatches(g, subMatches);
            subStandings[g] = calculateGroupStandings(g, fallbackOrder, subMatches);
        }
    });

    // Extract subFinalists and sub34 conditionally
    const actualFinalists = textResults && textResults.finalists.length ? textResults.finalists : [getTValue('M20') ? normalizeTeamName(String(getTValue('M20'))) : null, getTValue('M21') ? normalizeTeamName(String(getTValue('M21'))) : null].filter(Boolean);
    const actual34 = textResults && textResults.thirdFourth.length ? textResults.thirdFourth : [getTValue('M23') ? normalizeTeamName(String(getTValue('M23'))) : null, getTValue('M24') ? normalizeTeamName(String(getTValue('M24'))) : null].filter(Boolean);

    const subFinalists = upToMatchNr >= 104 ? actualFinalists : [];
    const sub34 = upToMatchNr >= 103 ? actual34 : [];

    return calculateLeaderboard(subMatches, subStandings, subFinalists, sub34, workbook);
}



function parseScorers(scorersVal) {
    if (!scorersVal || scorersVal === 'null' || scorersVal === 'undefined') return [];
    if (Array.isArray(scorersVal)) return scorersVal.map(cleanScorerName);
    if (typeof scorersVal === 'string') {
        let trimmed = scorersVal.trim();
        if (trimmed.startsWith('{') && trimmed.endsWith('}')) {
            trimmed = trimmed.slice(1, -1).trim();
        }
        if (trimmed.startsWith('[') && trimmed.endsWith(']')) {
            try {
                const parsed = JSON.parse(trimmed);
                if (Array.isArray(parsed)) return parsed.map(cleanScorerName);
            } catch (e) {
                // ignore
            }
        }
        return trimmed.split(',').map(s => cleanScorerName(s)).filter(Boolean);
    }
    return [];
}

function cleanScorerName(name) {
    if (!name) return '';
    let s = name.trim();
    // Remove all curly braces, smart quotes, and double quotes globally
    s = s.replace(/[\{\}“”"\u201C\u201D]/g, '');
    return s.trim();
}

function formatLiveStatus(timeElapsed) {
    if (!timeElapsed || timeElapsed === 'notstarted') return '';
    let elapsed = timeElapsed.trim().toUpperCase();
    if (elapsed === 'HT' || elapsed === 'HALF-TIME' || elapsed === 'HALFTIME') {
        return 'LIVE (HT)';
    }
    if (elapsed === 'FT' || elapsed === 'FINISHED') {
        return 'Päättynyt';
    }
    if (elapsed === 'LIVE') {
        return 'LIVE';
    }
    if (/^\d+$/.test(elapsed)) {
        return `LIVE ${elapsed}'`;
    }
    return `LIVE ${timeElapsed}`;
}

function isMatchFinishedByTime(dateStr, groupStr) {
    if (!dateStr) return false;
    const parts = dateStr.split(' klo ');
    if (parts.length < 2) return false;
    const dateParts = parts[0].split('.');
    if (dateParts.length < 2) return false;
    const day = parseInt(dateParts[0], 10);
    const month = parseInt(dateParts[1], 10); // 1-indexed for padding
    const hour = parseInt(parts[1], 10);

    const now = new Date();
    const currentYear = now.getFullYear();
    const year = currentYear >= 2026 ? currentYear : 2026;

    // Construct ISO string with Helsinki offset (+03:00) since World Cup is in June/July (Daylight Saving Time)
    const isoStr = `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}T${String(hour).padStart(2, '0')}:00:00+03:00`;
    const startDate = new Date(isoStr);
    
    if (isNaN(startDate.getTime())) return false;

    const elapsedMinutes = (now - startDate) / (1000 * 60);
    
    // 150 minutes (2.5h) for group stage, 210 minutes (3.5h) for playoffs
    const isGroup = !groupStr || groupStr.trim().length === 1;
    const threshold = isGroup ? 150 : 210;
    
    return elapsedMinutes > threshold;
}

async function fetchLiveApiData() {
    const cachePath = process.env.VERCEL 
        ? path.join('/tmp', 'cached-api-data.json') 
        : path.join(__dirname, 'cached-api-data.json');
        
    let cachedData = null;
    if (fs.existsSync(cachePath)) {
        try {
            cachedData = JSON.parse(fs.readFileSync(cachePath, 'utf8'));
        } catch (e) {
            console.warn('Failed to parse cached-api-data.json:', e.message);
        }
    }

    // Load existing live-games.json as a fallback/preserve cache
    let existingGames = [];
    const possiblePaths = [
        path.join(__dirname, '..', 'public', 'live-games.json'),
        path.join(process.cwd(), 'public', 'live-games.json'),
        '/var/task/public/live-games.json'
    ];
    for (const p of possiblePaths) {
        if (fs.existsSync(p)) {
            try {
                const data = JSON.parse(fs.readFileSync(p, 'utf8'));
                if (data && Array.isArray(data.games)) {
                    existingGames = data.games;
                    console.log(`[Cache Load] Loaded ${existingGames.length} existing games from ${p} as fallback/preserve cache.`);
                    break;
                }
            } catch (e) {}
        }
    }
    
    // If we don't have local cache path (e.g. running on Vercel but /tmp is empty),
    // let's try to load from the project directory's cached-api-data.json as a fallback read!
    if (!cachedData && process.env.VERCEL) {
        const localCachePath = path.join(__dirname, 'cached-api-data.json');
        if (fs.existsSync(localCachePath)) {
            try {
                cachedData = JSON.parse(fs.readFileSync(localCachePath, 'utf8'));
            } catch (e) {}
        }
    }

    let apiGames = [];
    let apiGroups = [];
    let apiTeams = [];

    // Load Excel sheet matches (1 to 72) as the source of truth for group stage team names
    const excelPath = path.join(__dirname, '..', 'MM2026_pistelaskenta.xlsx');
    
    const excelMatches = [];
    try {
        if (fs.existsSync(excelPath)) {
            const workbook = XLSX.readFile(excelPath);
            const tuloksetSheet = workbook.Sheets['Tulokset'];
            const getTValue = (addr) => tuloksetSheet[addr] ? tuloksetSheet[addr].v : undefined;
            for (let m = 1; m <= 72; m++) {
                const row = 5 + m;
                const t1 = getTValue(`D${row}`);
                const t2 = getTValue(`F${row}`);
                if (t1 && t2) {
                    excelMatches.push({
                        id: m,
                        team1: String(t1).trim(),
                        team2: String(t2).trim()
                    });
                }
            }
        }
    } catch (err) {
        console.warn('Could not read Excel for team names:', err.message);
    }

    const fetchOptions = {
        headers: {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
        }
    };

    async function fetchWithTimeout(url, options = {}, timeoutMs = 3000) {
        const controller = new AbortController();
        const id = setTimeout(() => controller.abort(), timeoutMs);
        try {
            const res = await fetch(url, {
                ...options,
                signal: controller.signal
            });
            clearTimeout(id);
            return res;
        } catch (err) {
            clearTimeout(id);
            throw err;
        }
    }

    async function fetchResource(url, proxyUrl) {
        try {
            const res = await fetchWithTimeout(url, fetchOptions, 3000);
            if (!res.ok) throw new Error(`HTTP ${res.status}`);
            return await res.json();
        } catch (err) {
            console.warn(`Direct fetch for ${url} failed. Trying proxy...`, err.message);
            try {
                const res = await fetchWithTimeout(proxyUrl, fetchOptions, 3000);
                if (!res.ok) throw new Error(`Proxy HTTP ${res.status}`);
                return await res.json();
            } catch (proxyErr) {
                console.error(`All fetches failed for ${url}:`, proxyErr.message);
                return null;
            }
        }
    }

    // Helper functions for football-data.org fallback
    function isAllowedFetchTime() {
        try {
            const options = { timeZone: 'Europe/Helsinki', hour: 'numeric', hour12: false };
            const formatter = new Intl.DateTimeFormat('en-US', options);
            const currentHour = parseInt(formatter.format(new Date()), 10);
            return currentHour >= 19 || currentHour < 9;
        } catch (e) {
            console.warn('Timezone Europe/Helsinki not supported, falling back to local time:', e.message);
            const currentHour = new Date().getHours();
            return currentHour >= 19 || currentHour < 9;
        }
    }

    function mapTeamToCacheName(fdName) {
        if (!fdName) return 'null';
        const norm = normalizeTeamName(fdName);
        if (norm === 'Tshekki') return 'Czech Republic';
        if (norm === 'Kap Verde') return 'Cape Verde';
        return fdName;
    }

    function formatDateString(utcDateStr) {
        if (!utcDateStr) return '';
        try {
            const d = new Date(utcDateStr);
            if (isNaN(d.getTime())) return utcDateStr;
            const mm = String(d.getMonth() + 1).padStart(2, '0');
            const dd = String(d.getDate()).padStart(2, '0');
            const yyyy = d.getFullYear();
            const hh = String(d.getHours()).padStart(2, '0');
            const min = String(d.getMinutes()).padStart(2, '0');
            return `${mm}/${dd}/${yyyy} ${hh}:${min}`;
        } catch (e) {
            return utcDateStr;
        }
    }

    // Try primary API (football-data.org) first
    let primarySuccess = false;

    // If running on GitHub Actions, try to fetch from the live Vercel deployment first
    // to leverage Vercel's geoblocking bypass and caching!
    if (process.env.GITHUB_ACTIONS === 'true') {
        try {
            console.log('[GitHub Actions] Fetching latest live games from production endpoint...');
            const res = await fetchWithTimeout('https://veikkaus2026.vercel.app/api/live-games', {}, 5000);
            if (res.ok) {
                const data = await res.json();
                if (data && Array.isArray(data.games) && data.games.length > 0) {
                    console.log(`[GitHub Actions] Successfully retrieved ${data.games.length} games from production.`);
                    apiGames = data.games;
                    apiGroups = cachedData ? cachedData.apiGroups : [];
                    apiTeams = cachedData ? cachedData.apiTeams : [];
                    primarySuccess = true;
                }
            } else {
                console.warn(`[GitHub Actions] Production endpoint returned HTTP ${res.status}`);
            }
        } catch (e) {
            console.warn('[GitHub Actions] Production endpoint fetch failed, falling back to direct APIs:', e.message);
        }
    }
    
    // Check 1-minute caching threshold for football-data.org to avoid rate limits
    const now = Date.now();
    let useCache = false;
    if (!primarySuccess && cachedData && cachedData.footballDataLastFetched) {
        const elapsedMs = now - cachedData.footballDataLastFetched;
        if (elapsedMs < 60000) {
            console.log(`[Cache] Using cached football-data.org results (fetched ${Math.round(elapsedMs / 1000)}s ago).`);
            useCache = true;
        }
    }

    if (useCache) {
        return cachedData;
    }

    // Retrieve token
    const token = process.env.FOOTBALL_DATA_API_KEY || process.env.FOOTBALL_DATA_TOKEN || process.env.FOOTBALL_DATA_API_TOKEN || "a07d5072049f468aa1425c0e32451068";
    if (!primarySuccess && token) {
        try {
            console.log('Fetching live matches from football-data.org...');
            const fdRes = await fetchWithTimeout('https://api.football-data.org/v4/competitions/WC/matches', {
                headers: {
                    'X-Auth-Token': token,
                    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
                }
            }, 3000);
            if (!fdRes.ok) {
                throw new Error(`football-data.org HTTP ${fdRes.status}`);
            }
            const fdData = await fdRes.json();
            if (!fdData || !fdData.matches) {
                throw new Error('Invalid response from football-data.org');
            }

            console.log(`Successfully fetched ${fdData.matches.length} matches from football-data.org. Mapping results...`);
            
            // Map the matches
            const mappedGames = [];
            const cachedGamesList = (cachedData && cachedData.apiGames) || [];
            
            // Pre-collect playoff matches from Football-Data.org to sort chronologically by stage
            const fdPlayoffsByStage = {};
            fdData.matches.forEach(m => {
                if (m.stage && m.stage !== 'GROUP_STAGE') {
                    if (!fdPlayoffsByStage[m.stage]) fdPlayoffsByStage[m.stage] = [];
                    fdPlayoffsByStage[m.stage].push(m);
                }
            });
            
            // Sort playoff matches in each stage chronologically
            Object.keys(fdPlayoffsByStage).forEach(stage => {
                fdPlayoffsByStage[stage].sort((a, b) => new Date(a.utcDate) - new Date(b.utcDate));
            });

            const PLAYOFF_CHRONO_MAP = {
                'LAST_32': [73, 76, 74, 75, 78, 77, 79, 80, 82, 81, 84, 83, 85, 88, 86, 87],
                'LAST_16': [90, 89, 91, 92, 93, 94, 95, 96],
                'QUARTER_FINALS': [97, 98, 99, 100],
                'SEMI_FINALS': [101, 102],
                'THIRD_PLACE': [103],
                'FINAL': [104]
            };

            // Process matches
            let hasStaleMatches = false;
            fdData.matches.forEach(m => {
                let matchId = null;
                
                if (m.stage === 'GROUP_STAGE') {
                    // Match by team names against Excel matches
                    const homeNorm = normalizeTeamName(m.homeTeam && m.homeTeam.name);
                    const awayNorm = normalizeTeamName(m.awayTeam && m.awayTeam.name);
                    
                    if (homeNorm && awayNorm) {
                        const foundGame = excelMatches.length > 0 
                            ? excelMatches.find(em => {
                                const emHomeNorm = normalizeTeamName(em.team1);
                                const emAwayNorm = normalizeTeamName(em.team2);
                                return (emHomeNorm === homeNorm && emAwayNorm === awayNorm) || 
                                       (emHomeNorm === awayNorm && emAwayNorm === homeNorm);
                              })
                            : cachedGamesList.find(cg => {
                                const cgHomeNorm = normalizeTeamName(cg.home_team_name_en);
                                const cgAwayNorm = normalizeTeamName(cg.away_team_name_en);
                                return (cgHomeNorm === homeNorm && cgAwayNorm === awayNorm) || 
                                       (cgHomeNorm === awayNorm && cgAwayNorm === homeNorm);
                              });
                              
                        if (foundGame) {
                            matchId = Number(foundGame.id);
                        }
                    }
                } else {
                    // Playoff match mapping
                    const stageList = fdPlayoffsByStage[m.stage] || [];
                    const index = stageList.indexOf(m);
                    const mapIds = PLAYOFF_CHRONO_MAP[m.stage];
                    if (index !== -1 && mapIds && mapIds[index]) {
                        matchId = mapIds[index];
                    }
                }
                
                if (matchId) {
                    const status = m.status; // FINISHED, IN_PLAY, PAUSED, SCHEDULED, TIMED
                    
                    // Check if match should have started but is still marked as TIMED/SCHEDULED
                    if ((status === 'TIMED' || status === 'SCHEDULED') && m.utcDate) {
                        const matchTime = new Date(m.utcDate);
                        // If current time is > matchTime + 15 minutes
                        if (Date.now() - matchTime.getTime() > 15 * 60 * 1000) {
                            console.log(`[Stale Check] Match ${matchId} should have started at ${m.utcDate} but status is still ${status}. Marking primary API as stale.`);
                            hasStaleMatches = true;
                        }
                    }

                    // Check if match should have finished but is still marked as live (IN_PLAY or PAUSED)
                    if ((status === 'IN_PLAY' || status === 'PAUSED') && m.utcDate) {
                        const matchTime = new Date(m.utcDate);
                        const isPlayoff = m.stage && m.stage !== 'GROUP_STAGE';
                        const durationThreshold = isPlayoff ? 200 * 60 * 1000 : 110 * 60 * 1000;
                        if (Date.now() - matchTime.getTime() > durationThreshold) {
                            console.log(`[Stale Check] Match ${matchId} should be finished (started at ${m.utcDate}, stage ${m.stage}) but status is still ${status}. Marking primary API as stale.`);
                            hasStaleMatches = true;
                        }
                    }

                    const isFinished = status === 'FINISHED';
                    const isLive = status === 'IN_PLAY' || status === 'PAUSED';
                    
                    let timeElapsed = 'notstarted';
                    if (isFinished) timeElapsed = 'finished';
                    else if (isLive) {
                        timeElapsed = m.minute ? String(m.minute) : 'live';
                    }
                    
                    const homeScore = m.score && m.score.fullTime && m.score.fullTime.home !== null 
                        ? String(m.score.fullTime.home) 
                        : '0';
                    const awayScore = m.score && m.score.fullTime && m.score.fullTime.away !== null 
                        ? String(m.score.fullTime.away) 
                        : '0';
                    
                    mappedGames.push({
                        id: String(matchId),
                        home_team_id: m.homeTeam ? String(m.homeTeam.id) : '0',
                        away_team_id: m.awayTeam ? String(m.awayTeam.id) : '0',
                        home_score: homeScore,
                        away_score: awayScore,
                        home_scorers: 'null',
                        away_scorers: 'null',
                        group: m.group ? m.group.replace('GROUP_', '') : '',
                        matchday: String(m.matchday || '1'),
                        local_date: formatDateString(m.utcDate),
                        finished: isFinished ? 'TRUE' : 'FALSE',
                        time_elapsed: timeElapsed,
                        type: m.stage ? m.stage.toLowerCase() : 'group',
                        home_team_name_en: m.homeTeam ? mapTeamToCacheName(m.homeTeam.name) : 'null',
                        away_team_name_en: m.awayTeam ? mapTeamToCacheName(m.awayTeam.name) : 'null'
                    });
                }
            });

            if (hasStaleMatches || mappedGames.length === 0) {
                console.warn('[Stale/Mapping Check] Primary API data has stale matches or mapped 0 games (stale: ' + hasStaleMatches + ', mapped: ' + mappedGames.length + '). Setting primarySuccess = false to trigger fallback to worldcup26.ir...');
                primarySuccess = false;
            } else {
                apiGames = mappedGames;
                apiGroups = cachedData ? cachedData.apiGroups : [];
                apiTeams = cachedData ? cachedData.apiTeams : [];
                primarySuccess = true;
            }
            
            // Save to cache
            try {
                fs.writeFileSync(cachePath, JSON.stringify({
                    footballDataLastFetched: Date.now(),
                    apiGames,
                    apiGroups,
                    apiTeams
                }, null, 2));
                console.log('Football-Data.org API data cached successfully to', cachePath);
            } catch (writeErr) {
                console.warn('Could not write cache (expected in read-only Vercel serverless):', writeErr.message);
            }
        } catch (fdErr) {
            console.error('Primary API (football-data.org) fetch failed:', fdErr.message);
        }
    }

    // Fall back to worldcup26.ir if football-data.org failed
    if (!primarySuccess) {
        console.log('Attempting fallback to worldcup26.ir API...');
        try {
            const { execSync } = require('child_process');
            let gamesData = await fetchResource(
                'https://worldcup26.ir/get/games',
                'https://api.codetabs.com/v1/proxy/?quest=https://worldcup26.ir/get/games'
            );
            
            // Local fallback using curl if fetch failed (useful on Windows)
            if ((!gamesData || !gamesData.games) && process.platform === 'win32') {
                try {
                    console.log('Node fetch failed. Trying curl.exe fallback for games...');
                    let stdout;
                    try {
                        stdout = execSync('curl.exe -k -s https://worldcup26.ir/get/games', { maxBuffer: 10 * 1024 * 1024 });
                    } catch (execErr) {
                        if (execErr.stdout && execErr.stdout.length > 0) {
                            stdout = execErr.stdout;
                        } else {
                            throw execErr;
                        }
                    }
                    gamesData = JSON.parse(stdout.toString());
                } catch (curlErr) {
                    console.warn('curl.exe fallback failed for games:', curlErr.message);
                }
            }

            if (!gamesData || !gamesData.games) {
                throw new Error('Failed to fetch live games from worldcup26.ir');
            }
            
            // Group playoff matches by type for chronological sorting
            const irPlayoffsByType = {};
            gamesData.games.forEach(g => {
                if (g.type && g.type !== 'group') {
                    if (!irPlayoffsByType[g.type]) irPlayoffsByType[g.type] = [];
                    irPlayoffsByType[g.type].push(g);
                }
            });
            
            // Sort each playoff type chronologically by local_date
            Object.keys(irPlayoffsByType).forEach(type => {
                irPlayoffsByType[type].sort((a, b) => new Date(a.local_date) - new Date(b.local_date));
            });

            const IR_PLAYOFF_TYPE_MAP = {
                'r32': 'LAST_32',
                'r16': 'LAST_16',
                'qf': 'QUARTER_FINALS',
                'sf': 'SEMI_FINALS',
                'third': 'THIRD_PLACE',
                'f': 'FINAL'
            };

            const PLAYOFF_CHRONO_MAP = {
                'LAST_32': [73, 76, 74, 75, 78, 77, 79, 80, 82, 81, 84, 83, 85, 88, 86, 87],
                'LAST_16': [90, 89, 91, 92, 93, 94, 95, 96],
                'QUARTER_FINALS': [97, 98, 99, 100],
                'SEMI_FINALS': [101, 102],
                'THIRD_PLACE': [103],
                'FINAL': [104]
            };

            // Map worldcup26.ir game IDs to Excel match IDs
            const mappedGames = [];
            gamesData.games.forEach(g => {
                let matchId = g.id;
                
                if (g.type === 'group') {
                    // Group stage mapping by team names
                    const homeNorm = normalizeTeamName(g.home_team_name_en);
                    const awayNorm = normalizeTeamName(g.away_team_name_en);
                    
                    if (homeNorm && awayNorm) {
                        const foundGame = excelMatches.find(em => {
                            const emHomeNorm = normalizeTeamName(em.team1);
                            const emAwayNorm = normalizeTeamName(em.team2);
                            return (emHomeNorm === homeNorm && emAwayNorm === awayNorm) || 
                                   (emHomeNorm === awayNorm && emAwayNorm === homeNorm);
                        });
                        if (foundGame) {
                            matchId = String(foundGame.id);
                        }
                    }
                } else {
                    // Playoff stage mapping chronologically
                    const mappedStage = IR_PLAYOFF_TYPE_MAP[g.type];
                    if (mappedStage) {
                        const typeList = irPlayoffsByType[g.type] || [];
                        const index = typeList.indexOf(g);
                        const mapIds = PLAYOFF_CHRONO_MAP[mappedStage];
                        if (index !== -1 && mapIds && mapIds[index]) {
                            matchId = String(mapIds[index]);
                        }
                    }
                }
                
                mappedGames.push({
                    ...g,
                    id: matchId
                });
            });

            apiGames = mappedGames;

            console.log('Fetching live groups from worldcup26.ir...');
            let groupsData = await fetchResource(
                'https://worldcup26.ir/get/groups',
                'https://api.codetabs.com/v1/proxy/?quest=https://worldcup26.ir/get/groups'
            );
            if ((!groupsData || !groupsData.groups) && process.platform === 'win32') {
                try {
                    console.log('Node fetch failed. Trying curl.exe fallback for groups...');
                    let stdout;
                    try {
                        stdout = execSync('curl.exe -k -s https://worldcup26.ir/get/groups', { maxBuffer: 10 * 1024 * 1024 });
                    } catch (execErr) {
                        if (execErr.stdout && execErr.stdout.length > 0) {
                            stdout = execErr.stdout;
                        } else {
                            throw execErr;
                        }
                    }
                    groupsData = JSON.parse(stdout.toString());
                } catch (e) {}
            }
            if (groupsData && groupsData.groups) {
                apiGroups = groupsData.groups;
            } else {
                apiGroups = cachedData ? cachedData.apiGroups : [];
            }

            console.log('Fetching live teams from worldcup26.ir...');
            let teamsData = await fetchResource(
                'https://worldcup26.ir/get/teams',
                'https://api.codetabs.com/v1/proxy/?quest=https://worldcup26.ir/get/teams'
            );
            if ((!teamsData || !teamsData.teams) && process.platform === 'win32') {
                try {
                    console.log('Node fetch failed. Trying curl.exe fallback for teams...');
                    let stdout;
                    try {
                        stdout = execSync('curl.exe -k -s https://worldcup26.ir/get/teams', { maxBuffer: 10 * 1024 * 1024 });
                    } catch (execErr) {
                        if (execErr.stdout && execErr.stdout.length > 0) {
                            stdout = execErr.stdout;
                        } else {
                            throw execErr;
                        }
                    }
                    teamsData = JSON.parse(stdout.toString());
                } catch (e) {}
            }
            if (teamsData && teamsData.teams) {
                apiTeams = teamsData.teams;
            } else {
                apiTeams = cachedData ? cachedData.apiTeams : [];
            }

            // Save to cache
            try {
                fs.writeFileSync(cachePath, JSON.stringify({ apiGames, apiGroups, apiTeams }, null, 2));
                console.log('Live API data cached successfully to', cachePath);
            } catch (writeErr) {
                console.warn('Could not write cache (expected in read-only Vercel serverless):', writeErr.message);
            }
        } catch (err) {
            console.error('Fallback API (worldcup26.ir) fetch failed:', err.message);
            if (cachedData && cachedData.apiGames) {
                console.log('Successfully loaded cached API data.');
                return cachedData;
            } else {
                throw new Error('Failed to fetch live API data and no valid cache exists: ' + err.message);
            }
        }
    }

    // Post-process apiGames to merge/preserve with existingGames
    if (existingGames.length > 0 && apiGames && apiGames.length > 0) {
        // 1. For any game in apiGames, if it's finished but has 0-0, check if existingGames has a non-zero score
        apiGames.forEach(g => {
            if (g.finished === 'TRUE' && g.home_score === '0' && g.away_score === '0') {
                const eg = existingGames.find(x => x.id === g.id);
                if (eg && eg.finished === 'TRUE' && (eg.home_score !== '0' || eg.away_score !== '0')) {
                    console.log(`[Preserve Score] Match ${g.id} was reset to 0-0 but was finished as ${eg.home_score}-${eg.away_score} in existing live-games.json. Restoring score.`);
                    g.home_score = eg.home_score;
                    g.away_score = eg.away_score;
                }
            }
        });

        // 2. If any game in existingGames is finished but is completely missing from apiGames, add it back!
        existingGames.forEach(eg => {
            if (eg.finished === 'TRUE' && !apiGames.some(x => x.id === eg.id)) {
                console.log(`[Preserve Match] Match ${eg.id} is missing from new API games but was finished in existing live-games.json. Adding it back.`);
                apiGames.push(eg);
            }
        });
    }

    return { apiGames, apiGroups, apiTeams };
}


async function runAnalysisGenerator({ excelPath, txtPath, apiGames = [], apiGroups = [], apiTeams = [], outputDir, writeFiles = true }) {
    const outputPath = path.join(outputDir || 'public', 'agent-analysis.json');

    // 1. Read files
    if (!fs.existsSync(excelPath)) {
        throw new Error(`Excel file ${excelPath} not found.`);
    }

    let textResults = null;
    if (fs.existsSync(txtPath)) {
        const text = fs.readFileSync(txtPath, 'utf8');
        textResults = parseResultsText(text);
    }

    const workbook = XLSX.readFile(excelPath);
    const tuloksetSheet = workbook.Sheets['Tulokset'];
    const getTValue = (addr) => tuloksetSheet[addr] ? tuloksetSheet[addr].v : undefined;

    // 2. Parse matches 1 to 72
    const matches = [];
    let playedCount = 0;
    for (let m = 1; m <= 72; m++) {
        const row = 5 + m;
        const nr = getTValue(`A${row}`);
        const date = getTValue(`B${row}`);
        const group = getTValue(`C${row}`);
        const t1 = getTValue(`D${row}`);
        const t2 = getTValue(`F${row}`);
        
        let g1 = getTValue(`G${row}`);
        let g2 = getTValue(`I${row}`);
        let hasResult = g1 !== undefined && g1 !== null && g1 !== '' && g2 !== undefined && g2 !== null && g2 !== '';
        let status = hasResult ? 'Päättynyt' : 'Pelaamaton';

        let homeScorers = null;
        let awayScorers = null;
        let apiHomeScore = null;
        let apiAwayScore = null;

        const apiGame = apiGames.find(g => Number(g.id) === m);
        if (apiGame) {
            homeScorers = apiGame.home_scorers;
            awayScorers = apiGame.away_scorers;
            apiHomeScore = apiGame.home_score !== null && apiGame.home_score !== undefined ? Number(apiGame.home_score) : null;
            apiAwayScore = apiGame.away_score !== null && apiGame.away_score !== undefined ? Number(apiGame.away_score) : null;
        }

        if (apiGame && apiGame.time_elapsed !== 'notstarted') {
            g1 = Number(apiGame.home_score);
            g2 = Number(apiGame.away_score);
            hasResult = true;
            if (apiGame.finished === 'TRUE' || isMatchFinishedByTime(date, group)) {
                status = 'Päättynyt';
            } else {
                status = formatLiveStatus(apiGame.time_elapsed);
            }
        }
        
        if (textResults && textResults.matches && textResults.matches[m]) {
            g1 = textResults.matches[m].goals1;
            g2 = textResults.matches[m].goals2;
            hasResult = true;
            status = 'Päättynyt';
        }

        if (hasResult) playedCount++;

        const showScorers = apiGame && hasResult && 
                            apiHomeScore !== null && apiAwayScore !== null && 
                            (Number(g1) === apiHomeScore) && (Number(g2) === apiAwayScore);

        matches.push({
            nr: Number(nr),
            date: String(date || ''),
            group: String(group || ''),
            team1: String(t1 || ''),
            team2: String(t2 || ''),
            goals1: hasResult ? Number(g1) : null,
            goals2: hasResult ? Number(g2) : null,
            hasResult: hasResult,
            status: status,
            originallyHadResult: hasResult,
            homeScorers: showScorers ? homeScorers : null,
            awayScorers: showScorers ? awayScorers : null
        });
    }

    // 3. Compute Standings & Playoff status
    const actualStandings = {};
    const groups = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L'];
    groups.forEach((g, gIdx) => {
        if (textResults && textResults.standings && textResults.standings[g]) {
            actualStandings[g] = textResults.standings[g];
        } else {
            let apiOrder = null;
            const apiGroup = apiGroups.find(x => x.name === g);
            if (apiGroup && apiGroup.teams && apiTeams.length > 0) {
                const mappedOrder = apiGroup.teams.map(tInfo => {
                    const teamObj = apiTeams.find(t => t.id === tInfo.team_id);
                    return teamObj ? normalizeTeamName(teamObj.name_en) : null;
                }).filter(Boolean);
                if (mappedOrder.length === 4) {
                    apiOrder = mappedOrder;
                }
            }

            if (apiOrder) {
                actualStandings[g] = apiOrder;
            } else {
                const isSecondHalf = gIdx >= 6;
                const startRow = isSecondHalf ? 12 : 6;
                const colLetter = XLSX.utils.encode_col(12 + (gIdx % 6));
                const originalOrder = [];
                for (let pos = 1; pos <= 4; pos++) {
                    const val = getTValue(`${colLetter}${startRow + pos - 1}`);
                    originalOrder.push(val ? normalizeTeamName(String(val)) : null);
                }
                const cleanOrder = originalOrder.filter(Boolean);
                const fallbackOrder = cleanOrder.length === 4 ? cleanOrder : getGroupTeamsFromMatches(g, matches);
                actualStandings[g] = calculateGroupStandings(g, fallbackOrder, matches);
            }
        }
    });

    let actualFinalists = [];
    if (textResults && textResults.finalists.length) {
        actualFinalists = textResults.finalists;
    } else {
        const apiFinalGame = apiGames.find(g => Number(g.id) === 104);
        if (apiFinalGame && apiFinalGame.home_team_name_en && apiFinalGame.away_team_name_en && apiFinalGame.home_team_name_en !== 'null' && apiFinalGame.away_team_name_en !== 'null') {
            actualFinalists = [
                normalizeTeamName(apiFinalGame.home_team_name_en),
                normalizeTeamName(apiFinalGame.away_team_name_en)
            ].filter(Boolean);
        } else {
            actualFinalists = [getTValue('M20') ? normalizeTeamName(String(getTValue('M20'))) : null, getTValue('M21') ? normalizeTeamName(String(getTValue('M21'))) : null].filter(Boolean);
        }
    }

    let actual34 = [];
    if (textResults && textResults.thirdFourth.length) {
        actual34 = textResults.thirdFourth;
    } else {
        const apiThirdPlaceGame = apiGames.find(g => Number(g.id) === 103);
        if (apiThirdPlaceGame && apiThirdPlaceGame.home_team_name_en && apiThirdPlaceGame.away_team_name_en && apiThirdPlaceGame.home_team_name_en !== 'null' && apiThirdPlaceGame.away_team_name_en !== 'null') {
            actual34 = [
                normalizeTeamName(apiThirdPlaceGame.home_team_name_en),
                normalizeTeamName(apiThirdPlaceGame.away_team_name_en)
            ].filter(Boolean);
        } else {
            actual34 = [getTValue('M23') ? normalizeTeamName(String(getTValue('M23'))) : null, getTValue('M24') ? normalizeTeamName(String(getTValue('M24'))) : null].filter(Boolean);
        }
    }

    const currentLeaderboard = calculateLeaderboard(matches, actualStandings, actualFinalists, actual34, workbook);

    // 5. Calculate Windows (V2 schema)
    const playedMatches = matches.filter(m => m.hasResult).sort((a, b) => parseMatchDate(a.date) - parseMatchDate(b.date) || a.nr - b.nr);
    let lastMatchWindow = { matches: [], participantPoints: [] };
    let last3MatchesWindow = { matches: [], participantPoints: [] };
    let last5MatchesWindow = { matches: [], participantPoints: [] };
    let latestDateWindow = { matches: [], participantPoints: [] };

    let summaryLastMatchNr = 0;
    let summaryLatestDate = "";

    const buildV2Window = (windowSize) => {
        if (playedMatches.length === 0) {
            return {
                matches: [],
                participantPoints: []
            };
        }

        const startIdx = Math.max(0, playedMatches.length - windowSize);
        const subMatchesInWindow = playedMatches.slice(startIdx);
        const minMatchNr = Math.min(...subMatchesInWindow.map(m => m.nr));
        const leaderboardBefore = getLeaderboardAtState(minMatchNr - 1, matches, textResults, workbook);

        const matchesSummary = subMatchesInWindow.map(m => ({
            nr: m.nr,
            date: m.date,
            group: m.group,
            team1: m.team1,
            team2: m.team2,
            goals1: m.goals1,
            goals2: m.goals2,
            resultText: `${m.goals1}-${m.goals2}`,
            hasResult: m.hasResult
        }));

        const participantPoints = currentLeaderboard.map(nowP => {
            const beforeP = leaderboardBefore.find(x => x.sheetName === nowP.sheetName) || { rank: currentLeaderboard.length, totalPoints: 0 };
            
            let pointsInWindow = 0;
            let exactResultsInWindow = 0;
            const matchPoints = [];

            subMatchesInWindow.forEach(m => {
                const pm = nowP.matches[m.nr - 1];
                let pts = null;
                let predStr = null;
                let exact = false;

                if (pm) {
                    pts = pm.points;
                    if (pm.hasGuess) {
                        predStr = `${pm.guess1}-${pm.guess2}`;
                        exact = (pts === 6);
                    }
                    if (pts !== null) {
                        pointsInWindow += pts;
                        if (pts === 6) exactResultsInWindow++;
                    }
                }

                matchPoints.push({
                    matchNr: m.nr,
                    team1: m.team1,
                    team2: m.team2,
                    resultText: `${m.goals1}-${m.goals2}`,
                    predictedResultText: predStr,
                    points: pts,
                    exactResult: exact
                });
            });

            return {
                name: nowP.name,
                rankNow: nowP.rank,
                totalPointsNow: nowP.totalPoints,
                pointsInWindow: pointsInWindow,
                rankBeforeWindow: beforeP.rank,
                rankChange: beforeP.rank - nowP.rank,
                exactResultsInWindow: exactResultsInWindow,
                matchPoints: matchPoints
            };
        });

        // Sort by pointsInWindow desc, then exactResultsInWindow desc, then rankNow asc
        participantPoints.sort((a, b) => b.pointsInWindow - a.pointsInWindow || b.exactResultsInWindow - a.exactResultsInWindow || a.rankNow - b.rankNow);

        return {
            matches: matchesSummary,
            participantPoints: participantPoints
        };
    };

    const buildV2WindowForDate = (dateMatches) => {
        if (dateMatches.length === 0) return { matches: [], participantPoints: [] };
        
        const minMatchNr = Math.min(...dateMatches.map(m => m.nr));
        const leaderboardBefore = getLeaderboardAtState(minMatchNr - 1, matches, textResults, workbook);

        const matchesSummary = dateMatches.map(m => ({
            nr: m.nr,
            date: m.date,
            group: m.group,
            team1: m.team1,
            team2: m.team2,
            goals1: m.goals1,
            goals2: m.goals2,
            resultText: `${m.goals1}-${m.goals2}`,
            hasResult: m.hasResult
        }));

        const participantPoints = currentLeaderboard.map(nowP => {
            const beforeP = leaderboardBefore.find(x => x.sheetName === nowP.sheetName) || { rank: currentLeaderboard.length, totalPoints: 0 };
            
            let pointsInWindow = 0;
            let exactResultsInWindow = 0;
            const matchPoints = [];

            dateMatches.forEach(m => {
                const pm = nowP.matches[m.nr - 1];
                let pts = null;
                let predStr = null;
                let exact = false;

                if (pm) {
                    pts = pm.points;
                    if (pm.hasGuess) {
                        predStr = `${pm.guess1}-${pm.guess2}`;
                        exact = (pts === 6);
                    }
                    if (pts !== null) {
                        pointsInWindow += pts;
                        if (pts === 6) exactResultsInWindow++;
                    }
                }

                matchPoints.push({
                    matchNr: m.nr,
                    team1: m.team1,
                    team2: m.team2,
                    resultText: `${m.goals1}-${m.goals2}`,
                    predictedResultText: predStr,
                    points: pts,
                    exactResult: exact
                });
            });

            return {
                name: nowP.name,
                rankNow: nowP.rank,
                totalPointsNow: nowP.totalPoints,
                pointsInWindow: pointsInWindow,
                rankBeforeWindow: beforeP.rank,
                rankChange: beforeP.rank - nowP.rank,
                exactResultsInWindow: exactResultsInWindow,
                matchPoints: matchPoints
            };
        });

        participantPoints.sort((a, b) => b.pointsInWindow - a.pointsInWindow || b.exactResultsInWindow - a.exactResultsInWindow || a.rankNow - b.rankNow);

        return {
            matches: matchesSummary,
            participantPoints: participantPoints
        };
    };

    const reportingDays = [
        { date: "12.6.", matches: [1, 2] },
        { date: "13.6.", matches: [3, 4] },
        { date: "14.6.", matches: [5, 6, 7, 8] },
        { date: "15.6.", matches: [9, 10, 11, 12] },
        { date: "16.6.", matches: [13, 14, 15, 16] },
        { date: "17.6.", matches: [17, 18, 19, 20] },
        { date: "18.6.", matches: [21, 22, 23, 24] },
        { date: "19.6.", matches: [25, 26, 27, 28] },
        { date: "20.6.", matches: [29, 30, 31, 32] },
        { date: "21.6.", matches: [33, 34, 35, 36] },
        { date: "22.6.", matches: [37, 38, 39, 40] },
        { date: "23.6.", matches: [41, 42, 43, 44] },
        { date: "24.6.", matches: [45, 46, 47, 48] },
        { date: "25.6.", matches: [49, 50, 51, 52, 53, 54] },
        { date: "26.6.", matches: [55, 56, 57, 58, 59, 60] },
        { date: "27.6.", matches: [61, 62, 63, 64, 65, 66] },
        { date: "28.6.", matches: [67, 68, 69, 70, 71, 72] }
    ];

    let currentReportingDay = null;
    let reportingWindow = { matches: [], participantPoints: [] };

    const buildCustomMatchesWindow = (matchNrs) => {
        if (playedMatches.length === 0 || matchNrs.length === 0) {
            return {
                matches: [],
                participantPoints: []
            };
        }

        const subMatchesInWindow = playedMatches.filter(m => matchNrs.includes(m.nr));
        if (subMatchesInWindow.length === 0) {
            return {
                matches: [],
                participantPoints: []
            };
        }

        const minMatchNr = Math.min(...matchNrs);
        const leaderboardBefore = getLeaderboardAtState(minMatchNr - 1, matches, textResults, workbook);

        const matchesSummary = subMatchesInWindow.map(m => ({
            nr: m.nr,
            date: m.date,
            group: m.group,
            team1: m.team1,
            team2: m.team2,
            goals1: m.goals1,
            goals2: m.goals2,
            resultText: `${m.goals1}-${m.goals2}`,
            hasResult: m.hasResult
        }));

        const participantPoints = currentLeaderboard.map(nowP => {
            const beforeP = leaderboardBefore.find(x => x.sheetName === nowP.sheetName) || { rank: currentLeaderboard.length, totalPoints: 0 };
            
            let pointsInWindow = 0;
            let exactResultsInWindow = 0;
            const matchPoints = [];

            subMatchesInWindow.forEach(m => {
                const pm = nowP.matches[m.nr - 1];
                let pts = null;
                let predStr = null;
                let exact = false;

                if (pm) {
                    pts = pm.points;
                    if (pm.hasGuess) {
                        predStr = `${pm.guess1}-${pm.guess2}`;
                        exact = (pts === 6);
                    }
                    if (pts !== null) {
                        pointsInWindow += pts;
                        if (pts === 6) exactResultsInWindow++;
                    }
                }

                matchPoints.push({
                    matchNr: m.nr,
                    team1: m.team1,
                    team2: m.team2,
                    resultText: `${m.goals1}-${m.goals2}`,
                    predictedResultText: predStr,
                    points: pts,
                    exactResult: exact
                });
            });

            return {
                name: nowP.name,
                rankNow: nowP.rank,
                totalPointsNow: nowP.totalPoints,
                pointsInWindow: pointsInWindow,
                rankBeforeWindow: beforeP.rank,
                rankChange: beforeP.rank - nowP.rank,
                exactResultsInWindow: exactResultsInWindow,
                matchPoints: matchPoints
            };
        });

        participantPoints.sort((a, b) => b.pointsInWindow - a.pointsInWindow || b.exactResultsInWindow - a.exactResultsInWindow || a.rankNow - b.rankNow);

        return {
            matches: matchesSummary,
            participantPoints: participantPoints
        };
    };

    if (playedMatches.length > 0) {
        const L = playedMatches[playedMatches.length - 1];
        summaryLastMatchNr = L.nr;
        summaryLatestDate = L.date;

        lastMatchWindow = buildV2Window(1);
        last3MatchesWindow = buildV2Window(3);
        last5MatchesWindow = buildV2Window(5);

        let dateMatches = [];
        const latestDayConfig = reportingDays.find(d => d.matches.includes(summaryLastMatchNr));
        if (latestDayConfig) {
            dateMatches = playedMatches.filter(m => latestDayConfig.matches.includes(m.nr));
        } else {
            const latestDatePrefix = L.date.split(' ')[0];
            dateMatches = playedMatches.filter(m => m.date.startsWith(latestDatePrefix));
        }
        latestDateWindow = buildV2WindowForDate(dateMatches);

        currentReportingDay = reportingDays.find(d => d.matches.includes(summaryLastMatchNr));
        if (currentReportingDay) {
            reportingWindow = buildCustomMatchesWindow(currentReportingDay.matches);
        }
    }

    // 6. Build Detailed Participant Data
    const participantsData = currentLeaderboard.map(p => {
        let runningSum = 0;
        const participantMatches = p.matches.map(pm => {
            const actual = matches[pm.nr - 1];
            if (actual.hasResult && pm.points !== null) {
                runningSum += pm.points;
            }
            const winnerCorrect = (actual.hasResult && pm.hasGuess) 
                ? (Math.sign(pm.guess1 - pm.guess2) === Math.sign(actual.goals1 - actual.goals2)) 
                : null;
            const exactCorrect = (actual.hasResult && pm.hasGuess) 
                ? (pm.points === 6) 
                : null;

            return {
                matchNr: pm.nr,
                teams: { team1: actual.team1, team2: actual.team2 },
                actualScore: actual.hasResult ? { goals1: actual.goals1, goals2: actual.goals2 } : null,
                guessScore: pm.hasGuess ? { goals1: pm.guess1, goals2: pm.guess2 } : null,
                points: pm.points,
                cumulativePoints: actual.hasResult ? runningSum : null,
                isWinnerCorrect: winnerCorrect,
                isExactCorrect: exactCorrect
            };
        });

        // Compute last-N-matches points
        let lastMatchPoints = 0;
        let last3MatchesPoints = 0;
        let last5MatchesPoints = 0;
        let latestDatePoints = 0;

        if (playedMatches.length > 0) {
            const lastM = playedMatches[playedMatches.length - 1];
            const pmL = p.matches[lastM.nr - 1];
            if (pmL && pmL.points !== null) lastMatchPoints = pmL.points;

            const last3 = playedMatches.slice(-3);
            last3.forEach(m => {
                const pm = p.matches[m.nr - 1];
                if (pm && pm.points !== null) last3MatchesPoints += pm.points;
            });

            const last5 = playedMatches.slice(-5);
            last5.forEach(m => {
                const pm = p.matches[m.nr - 1];
                if (pm && pm.points !== null) last5MatchesPoints += pm.points;
            });

            const latestDatePrefix = lastM.date.split(' ')[0];
            const dateMatches = playedMatches.filter(m => m.date.startsWith(latestDatePrefix));
            dateMatches.forEach(m => {
                const pm = p.matches[m.nr - 1];
                if (pm && pm.points !== null) latestDatePoints += pm.points;
            });
        }

        return {
            name: p.name,
            sheetName: p.sheetName,
            rank: p.rank,
            totalPoints: p.totalPoints,
            matchPoints: p.matchPoints,
            groupPoints: p.groupPoints,
            finalPoints: p.finalPoints,
            exactlyCorrect: p.exactlyCorrect,
            groupPointsAdded: p.groupPointsAdded,
            finalistsAdded: p.finalistsAdded,
            thirdFourthAdded: p.thirdFourthAdded,
            // matches: participantMatches, // Omitted to keep JSON file size small (< 100 KB) for AI agent consumption
            aggregates: {
                pointsLastMatch: lastMatchPoints,
                pointsLast3Matches: last3MatchesPoints,
                pointsLast5Matches: last5MatchesPoints,
                pointsLatestDateMatches: latestDatePoints
            }
        };
    });

    // 7. Helper to build window analysis for Botmanen summary
    const buildWindowAnalysis = (windowSize) => {
        const w = buildV2Window(windowSize);
        const zeroPointers = w.participantPoints
            .filter(p => p.pointsInWindow === 0)
            .map(p => ({
                name: p.name,
                rankNow: p.rankNow,
                pointsInWindow: 0
            }));

        const topPerformers = w.participantPoints.slice(0, 5).map(p => ({
            name: p.name,
            pointsInWindow: p.pointsInWindow,
            exactResultsInWindow: p.exactResultsInWindow
        }));

        const perfectScores = [];
        w.participantPoints.forEach(p => {
            p.matchPoints.forEach(pm => {
                if (pm.points === 6) {
                    perfectScores.push({
                        name: p.name,
                        matchNr: pm.matchNr,
                        match: `${pm.team1}–${pm.team2}`,
                        guess: pm.predictedResultText,
                        actual: pm.resultText
                    });
                }
            });
        });

        const biggestRisers = [...w.participantPoints]
            .filter(p => p.rankChange > 0)
            .sort((a, b) => b.rankChange - a.rankChange || b.pointsInWindow - a.pointsInWindow)
            .slice(0, 5)
            .map(p => ({
                name: p.name,
                rankBeforeWindow: p.rankBeforeWindow,
                rankNow: p.rankNow,
                rankChange: p.rankChange,
                pointsInWindow: p.pointsInWindow
            }));

        const biggestFallers = [...w.participantPoints]
            .filter(p => p.rankChange < 0)
            .sort((a, b) => a.rankChange - b.rankChange || a.pointsInWindow - b.pointsInWindow)
            .slice(0, 5)
            .map(p => ({
                name: p.name,
                rankBeforeWindow: p.rankBeforeWindow,
                rankNow: p.rankNow,
                rankChange: p.rankChange,
                pointsInWindow: p.pointsInWindow
            }));

        return {
            matches: w.matches,
            participantPoints: w.participantPoints,
            topPerformers: topPerformers,
            perfectScores: perfectScores,
            biggestRisers: biggestRisers,
            biggestFallers: biggestFallers,
            zeroPointers: zeroPointers
        };
    };

    // 8. commentaryHints for original JSON (V2 structure)
    const topPerformersLastWindow = last3MatchesWindow.participantPoints
        .slice(0, 5)
        .map(p => ({
            name: p.name,
            pointsInWindow: p.pointsInWindow,
            rankNow: p.rankNow,
            totalPointsNow: p.totalPointsNow
        }));

    const biggestRisers = [...last3MatchesWindow.participantPoints]
        .filter(p => p.rankChange > 0)
        .sort((a, b) => b.rankChange - a.rankChange || b.pointsInWindow - a.pointsInWindow)
        .slice(0, 5)
        .map(p => ({
            name: p.name,
            rankChange: p.rankChange,
            rankBeforeWindow: p.rankBeforeWindow,
            rankNow: p.rankNow,
            pointsInWindow: p.pointsInWindow
        }));

    const biggestFallers = [...last3MatchesWindow.participantPoints]
        .filter(p => p.rankChange < 0)
        .sort((a, b) => a.rankChange - b.rankChange || a.pointsInWindow - b.pointsInWindow)
        .slice(0, 5)
        .map(p => ({
            name: p.name,
            rankChange: p.rankChange,
            rankBeforeWindow: p.rankBeforeWindow,
            rankNow: p.rankNow,
            pointsInWindow: p.pointsInWindow
        }));

    const perfectScores = last3MatchesWindow.participantPoints
        .filter(p => p.exactResultsInWindow > 0)
        .map(p => {
            const perfMatches = p.matchPoints
                .filter(mp => mp.exactResult)
                .map(mp => ({
                    matchNr: mp.matchNr,
                    team1: mp.team1,
                    team2: mp.team2,
                    resultText: mp.resultText
                }));
            return {
                name: p.name,
                exactResultsInWindow: p.exactResultsInWindow,
                matches: perfMatches
            };
        });

    const tightGroups = [];
    groups.forEach(g => {
        const groupMatches = matches.filter(m => m.group === g);
        const playedGroupMatches = groupMatches.filter(m => m.hasResult);
        if (playedGroupMatches.length > 0) {
            const stats = {};
            const teamList = actualStandings[g] || [];
            teamList.forEach(t => { stats[t] = 0; });
            groupMatches.forEach(m => {
                if (!m.hasResult) return;
                if (stats[m.team1] !== undefined && stats[m.team2] !== undefined) {
                    if (m.goals1 > m.goals2) stats[m.team1] += 3;
                    else if (m.goals1 < m.goals2) stats[m.team2] += 3;
                    else { stats[m.team1] += 1; stats[m.team2] += 1; }
                }
            });
            const pointsList = Object.values(stats).sort((a, b) => b - a);
            if (pointsList.length === 4) {
                const diff = pointsList[0] - pointsList[3];
                if (diff <= 3) {
                    tightGroups.push({ group: g, diff: diff, points: pointsList });
                }
            }
        }
    });

    const participantsNeedingComeback = currentLeaderboard
        .slice(-3)
        .map(p => ({ name: p.name, rank: p.rank, totalPoints: p.totalPoints }));

    const commentaryHints = {
        topPerformersLastWindow,
        biggestRisers,
        biggestFallers,
        perfectScores,
        tightGroups,
        participantsNeedingComeback
    };

    const emailLatestMatches = reportingWindow.matches || [];
    const emailTopPerformers = reportingWindow.participantPoints || [];
    const emailPerfectScores = (reportingWindow.participantPoints || [])
        .filter(p => p.exactResultsInWindow > 0)
        .map(p => {
            const perfMatches = p.matchPoints
                .filter(mp => mp.exactResult)
                .map(mp => ({
                    matchNr: mp.matchNr,
                    team1: mp.team1,
                    team2: mp.team2,
                    resultText: mp.resultText
                }));
            return {
                name: p.name,
                exactResultsInWindow: p.exactResultsInWindow,
                matches: perfMatches
            };
        });

    const emailLeaderboardTop = currentLeaderboard.slice(0, 12).map(p => {
        const winInfo = reportingWindow.participantPoints
            ? reportingWindow.participantPoints.find(x => x.name === p.name)
            : null;
        return {
            name: p.name,
            rankNow: p.rank,
            totalPointsNow: p.totalPoints,
            pointsInWindow: winInfo ? winInfo.pointsInWindow : 0
        };
    });

    const emailBiggestRisers = reportingWindow.participantPoints
        ? reportingWindow.participantPoints
            .filter(p => p.rankChange > 0)
            .sort((a, b) => b.rankChange - a.rankChange || b.pointsInWindow - a.pointsInWindow)
            .slice(0, 8)
            .map(p => ({
                name: p.name,
                rankChange: p.rankChange,
                rankBeforeWindow: p.rankBeforeWindow,
                rankNow: p.rankNow,
                pointsInWindow: p.pointsInWindow
            }))
        : [];

    const emailBiggestFallers = reportingWindow.participantPoints
        ? reportingWindow.participantPoints
            .filter(p => p.rankChange < 0)
            .sort((a, b) => a.rankChange - b.rankChange || a.pointsInWindow - b.pointsInWindow)
            .slice(0, 8)
            .map(p => ({
                name: p.name,
                rankChange: p.rankChange,
                rankBeforeWindow: p.rankBeforeWindow,
                rankNow: p.rankNow,
                pointsInWindow: p.pointsInWindow
            }))
        : [];

    const emailParticipantsNeedingComeback = reportingWindow.participantPoints
        ? reportingWindow.participantPoints
            .filter(p => p.pointsInWindow <= 3 || p.rankChange < 0)
            .sort((a, b) => a.pointsInWindow - b.pointsInWindow || a.rankChange - b.rankChange)
            .slice(0, 8)
            .map(p => ({
                name: p.name,
                rankNow: p.rankNow,
                rankChange: p.rankChange,
                pointsInWindow: p.pointsInWindow,
                totalPointsNow: p.totalPointsNow
            }))
        : [];

    const emailDigest = {
        window: currentReportingDay ? currentReportingDay.date : "last3Matches",
        generatedAt: new Date().toISOString(),
        summary: {
            playedMatchesCount: playedCount,
            totalMatches: 72,
            lastPlayedMatchNr: summaryLastMatchNr,
            latestPlayedDate: summaryLatestDate,
            participantCount: currentLeaderboard.length
        },
        latestMatches: emailLatestMatches,
        topPerformers: emailTopPerformers,
        perfectScores: emailPerfectScores,
        leaderboardTop: emailLeaderboardTop,
        biggestRisers: emailBiggestRisers,
        biggestFallers: emailBiggestFallers,
        participantsNeedingComeback: emailParticipantsNeedingComeback
    };

    // 9. Construct original agent-analysis.json
    const finalJson = {
        emailDigest: emailDigest,
        schemaVersion: 1,
        generatedAt: new Date().toISOString(),
        source: {
            excelFile: excelPath,
            resultsFile: txtPath,
            resultsPriority: "tulokset.txt overrides Excel match results when present"
        },
        summary: {
            playedMatchesCount: playedCount,
            totalMatches: 72,
            lastPlayedMatchNr: summaryLastMatchNr,
            latestPlayedDate: summaryLatestDate,
            participantCount: currentLeaderboard.length
        },
        matches: matches,
        leaderboard: currentLeaderboard.map(p => ({
            rank: p.rank,
            name: p.name,
            sheetName: p.sheetName,
            totalPoints: p.totalPoints,
            matchPoints: p.matchPoints,
            groupPoints: p.groupPoints,
            finalPoints: p.finalPoints,
            exactlyCorrect: p.exactlyCorrect,
            groupPointsAdded: p.groupPointsAdded,
            finalistsAdded: p.finalistsAdded,
            thirdFourthAdded: p.thirdFourthAdded
        })),
        participants: participantsData,
        latestWindows: {
            lastMatch: lastMatchWindow,
            last3Matches: last3MatchesWindow,
            last5Matches: last5MatchesWindow,
            latestDate: latestDateWindow
        },
        commentaryHints: commentaryHints
    };

    // 10. Construct new botmanen-summary.json
    const latestPlayedMatches = [];
    if (playedMatches.length > 0) {
        const L = playedMatches[playedMatches.length - 1];
        const latestDayConfig = reportingDays.find(d => d.matches.includes(L.nr));
        
        let filteredMatches = [];
        if (latestDayConfig) {
            filteredMatches = playedMatches.filter(m => latestDayConfig.matches.includes(m.nr));
        } else {
            const latestDatePrefix = L.date.split(' ')[0];
            filteredMatches = playedMatches.filter(m => m.date.startsWith(latestDatePrefix));
        }
        
        filteredMatches.forEach(m => {
            latestPlayedMatches.push({
                nr: m.nr,
                date: m.date,
                group: m.group,
                team1: m.team1,
                team2: m.team2,
                goals1: m.goals1,
                goals2: m.goals2
            });
        });
    }

    const leaderboardTop = currentLeaderboard.slice(0, 10).map(p => ({
        rank: p.rank,
        name: p.name,
        totalPoints: p.totalPoints,
        matchPoints: p.matchPoints,
        exactlyCorrect: p.exactlyCorrect
    }));

    const last3Analysis = buildWindowAnalysis(3);
    const last5Raw = buildWindowAnalysis(5);
    const last5Analysis = {
        participantPoints: last5Raw.participantPoints,
        topPerformers: last5Raw.topPerformers
    };

    const leader = currentLeaderboard.length > 0 ? {
        name: currentLeaderboard[0].name,
        points: currentLeaderboard[0].totalPoints
    } : { name: "", points: 0 };

    const closestChasers = currentLeaderboard
        .slice(1, 6)
        .map(p => ({
            name: p.name,
            rank: p.rank,
            points: p.totalPoints,
            pointsBehindLeader: leader.points - p.totalPoints
        }));

    let tightTopComment = "Kärjessä ei ole merkittävää pistekehitystä.";
    if (closestChasers.length > 0) {
        const diff = closestChasers[0].pointsBehindLeader;
        if (diff === 0) {
            tightTopComment = "Kärjessä on tasapeli!";
        } else if (diff === 1) {
            tightTopComment = "Kärki on yhden pisteen sisällä.";
        } else if (diff <= 3) {
            tightTopComment = `Kärki on erittäin tiukka, eroa vain ${diff} pistettä.`;
        } else {
            tightTopComment = `Johtaja pitää ${diff} pisteen eron seuraavaan.`;
        }
    }

    const suggestParts = [];
    if (currentLeaderboard.length > 0) {
        suggestParts.push(`Koko kisan kärjessä on ${leader.name} (${leader.points}p).`);
        if (closestChasers.length > 0) {
            const firstChaser = closestChasers[0];
            suggestParts.push(`Lähimpänä haastajana on ${firstChaser.name} (${firstChaser.points}p, eroa vain ${firstChaser.pointsBehindLeader}p).`);
        }
        if (last3Analysis.topPerformers.length > 0) {
            const tp = last3Analysis.topPerformers[0];
            suggestParts.push(`Viimeisen 3 ottelun aikana parasta tahtia on pitänyt ${tp.name} keräten ${tp.pointsInWindow}p.`);
        }
        if (last3Analysis.biggestRisers.length > 0) {
            const riser = last3Analysis.biggestRisers[0];
            suggestParts.push(`Suurin nousija sarjataulukossa viimeisen 3 ottelun aikana on ollut ${riser.name} (+${riser.rankChange} sijaa).`);
        }
        if (last3Analysis.biggestFallers.length > 0) {
            const faller = last3Analysis.biggestFallers[0];
            suggestParts.push(`Eniten sijoitustaan menetti ${faller.name} (${faller.rankChange} sijaa).`);
        }
        if (last3Analysis.perfectScores.length > 0) {
            const exacts = last3Analysis.perfectScores.slice(0, 3).map(ps => `${ps.name} (${ps.match} veikkaus ${ps.guess})`);
            suggestParts.push(`Täydellisiä 6 pisteen osumia saivat: ${exacts.join(', ')}.`);
        }
    }

    const botmanenJson = {
        schemaVersion: 2,
        generatedAt: new Date().toISOString(),
        summary: {
            playedMatchesCount: playedCount,
            totalMatches: 72,
            lastPlayedMatchNr: summaryLastMatchNr,
            latestPlayedDate: summaryLatestDate,
            participantCount: currentLeaderboard.length
        },
        latestPlayedMatches: latestPlayedMatches,
        leaderboardTop: leaderboardTop,
        last3MatchesAnalysis: last3Analysis,
        last5MatchesAnalysis: last5Analysis,
        commentaryHints: {
            leader: leader,
            closestChasers: closestChasers,
            tightTopComment: tightTopComment,
            suggestedFinnishSummary: suggestParts.join(" ")
        }
    };
    if (writeFiles && outputDir) {
        if (!fs.existsSync(outputDir)) {
            fs.mkdirSync(outputDir, { recursive: true });
        }
        fs.writeFileSync(outputPath, JSON.stringify(finalJson, null, 2));
        const botmanenOutputPath = path.join(outputDir, 'botmanen-summary.json');
        fs.writeFileSync(botmanenOutputPath, JSON.stringify(botmanenJson, null, 2));

        // Write HTML wrapper for agent analysis
        const htmlOutputPath = path.join(outputDir, 'agent-analysis-html.html');
        const htmlContent = `<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <title>Veikkaus 2026 Agent Data</title>
</head>
<body>
    <pre id="agent-data" style="word-wrap: break-word; white-space: pre-wrap;">${JSON.stringify(finalJson)}</pre>
</body>
</html>`;
        fs.writeFileSync(htmlOutputPath, htmlContent);

        // Write static live-games.json
        const liveGamesOutputPath = path.join(outputDir, 'live-games.json');
        fs.writeFileSync(liveGamesOutputPath, JSON.stringify({ games: apiGames }, null, 2));

        console.log(`\n=== Agent Analysis Generation Successful ===`);
        console.log(`Matches read: 72`);
        console.log(`Played matches found: ${playedCount}`);
        console.log(`Participants found: ${currentLeaderboard.length}`);
        console.log(`Output written to: ${outputPath}`);
        console.log(`Output written to: ${botmanenOutputPath}`);
        console.log(`Output written to: ${htmlOutputPath}`);
        console.log(`Output written to: ${liveGamesOutputPath}\n`);
    }

    return { finalJson, botmanenJson, playedCount, currentLeaderboardLength: currentLeaderboard.length };
}

if (require.main === module) {
    (async () => {
        try {
            console.log('Fetching live API data for static build...');
            const { apiGames, apiGroups, apiTeams } = await fetchLiveApiData();
            await runAnalysisGenerator({
                excelPath: 'MM2026_pistelaskenta.xlsx',
                txtPath: 'tulokset.txt',
                apiGames,
                apiGroups,
                apiTeams,
                outputDir: 'public',
                writeFiles: true
            });
        } catch (e) {
            console.error('Failed to run static generator:', e);
            if (process.env.GITHUB_ACTIONS === 'true') {
                console.log('Running on GitHub Actions: exiting gracefully with 0 to prevent workflow failure emails.');
                process.exit(0);
            } else {
                process.exit(1);
            }
        }
    })();
}

module.exports = {
    runAnalysisGenerator,
    normalizeTeamName,
    parseResultsText,
    fetchLiveApiData
};
