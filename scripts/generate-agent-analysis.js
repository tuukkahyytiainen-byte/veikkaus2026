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
    "england": "Englanti",
    "czech republic": "Tshekki",
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
    "united states": "USA",
    "australia": "Australia",
    "belgium": "Belgia",
    "iran": "Iran",
    "croatia": "Kroatia"
};

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

// MAIN EXECUTION
try {
    const excelPath = 'MM2026_pistelaskenta.xlsx';
    const txtPath = 'tulokset.txt';
    const outputDir = 'public';
    const outputPath = path.join(outputDir, 'agent-analysis.json');

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
        
        if (textResults && textResults.matches && textResults.matches[m]) {
            g1 = textResults.matches[m].goals1;
            g2 = textResults.matches[m].goals2;
            hasResult = true;
            status = 'Päättynyt';
        }
        
        if (hasResult) playedCount++;
        
        matches.push({
            nr: Number(nr),
            date: String(date || ''),
            group: String(group || ''),
            team1: String(t1 || ''),
            team2: String(t2 || ''),
            goals1: hasResult ? Number(g1) : null,
            goals2: hasResult ? Number(g2) : null,
            hasResult: hasResult,
            status: status
        });
    }

    // 3. Compute Standings & Playoff status
    const actualStandings = {};
    const groups = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L'];
    groups.forEach((g, gIdx) => {
        if (textResults && textResults.standings && textResults.standings[g]) {
            actualStandings[g] = textResults.standings[g];
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
    });

    const actualFinalists = textResults && textResults.finalists.length ? textResults.finalists : [getTValue('M20') ? normalizeTeamName(String(getTValue('M20'))) : null, getTValue('M21') ? normalizeTeamName(String(getTValue('M21'))) : null].filter(Boolean);
    const actual34 = textResults && textResults.thirdFourth.length ? textResults.thirdFourth : [getTValue('M23') ? normalizeTeamName(String(getTValue('M23'))) : null, getTValue('M24') ? normalizeTeamName(String(getTValue('M24'))) : null].filter(Boolean);

    // 4. Calculate Current Leaderboard
    const currentLeaderboard = calculateLeaderboard(matches, actualStandings, actualFinalists, actual34, workbook);

    // 5. Calculate Windows (V2 schema)
    const playedMatches = matches.filter(m => m.hasResult).sort((a, b) => a.nr - b.nr);
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
        const startMatch = subMatchesInWindow[0];
        const leaderboardBefore = getLeaderboardAtState(startMatch.nr - 1, matches, textResults, workbook);

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
        
        const startMatch = dateMatches[0];
        const leaderboardBefore = getLeaderboardAtState(startMatch.nr - 1, matches, textResults, workbook);

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

        const latestDatePrefix = L.date.split(' ')[0];
        const dateMatches = playedMatches.filter(m => m.date.startsWith(latestDatePrefix));
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
        const latestDatePrefix = L.date.split(' ')[0];
        playedMatches
            .filter(m => m.date.startsWith(latestDatePrefix))
            .forEach(m => {
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

    // Ensure output dir exists
    if (!fs.existsSync(outputDir)) {
        fs.mkdirSync(outputDir, { recursive: true });
    }

    fs.writeFileSync(outputPath, JSON.stringify(finalJson, null, 2));
    const botmanenOutputPath = path.join(outputDir, 'botmanen-summary.json');
    fs.writeFileSync(botmanenOutputPath, JSON.stringify(botmanenJson, null, 2));

    // Console logs (requirement 9)
    console.log(`\n=== Agent Analysis Generation Successful ===`);
    console.log(`Matches read: 72`);
    console.log(`Played matches found: ${playedCount}`);
    console.log(`Participants found: ${currentLeaderboard.length}`);
    console.log(`Output written to: ${outputPath}`);
    console.log(`Output written to: ${botmanenOutputPath}\n`);

} catch (err) {
    console.error('Error generating agent analysis:', err);
    process.exit(1);
}
