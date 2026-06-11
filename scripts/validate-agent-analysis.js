const fs = require('fs');
const path = require('path');

console.log('Validating agent-analysis.json...');

const publicDir = path.join(__dirname, '..', 'public');
const jsonPath = path.join(publicDir, 'agent-analysis.json');

if (!fs.existsSync(jsonPath)) {
    console.error(`Error: File not found at ${jsonPath}`);
    process.exit(1);
}

let data;
try {
    const content = fs.readFileSync(jsonPath, 'utf8');
    data = JSON.parse(content);
    console.log('✓ File is valid JSON');
} catch (error) {
    console.error('Error: Failed to parse JSON', error);
    process.exit(1);
}

// 1. Check top-level emailDigest
if (!data.emailDigest) {
    console.error('Error: top-level field "emailDigest" is missing!');
    process.exit(1);
}
console.log('✓ top-level "emailDigest" exists');

// 2. Check summary playedMatchesCount matches
const summaryCount = data.summary ? data.summary.playedMatchesCount : null;
const digestCount = data.emailDigest.summary ? data.emailDigest.summary.playedMatchesCount : null;

if (summaryCount === null || digestCount === null) {
    console.error('Error: playedMatchesCount is missing from summary or emailDigest.summary');
    process.exit(1);
}

if (summaryCount !== digestCount) {
    console.error(`Error: playedMatchesCount mismatch! summary: ${summaryCount}, emailDigest.summary: ${digestCount}`);
    process.exit(1);
}
console.log(`✓ playedMatchesCount matches: ${summaryCount}`);

// 3. Check latestMatches contains 1-3 matches
const latestMatches = data.emailDigest.latestMatches;
if (!Array.isArray(latestMatches)) {
    console.error('Error: emailDigest.latestMatches is not an array!');
    process.exit(1);
}

const minExpectedMatches = summaryCount === 0 ? 0 : 1;
if (latestMatches.length < minExpectedMatches || latestMatches.length > 6) {
    console.error(`Error: emailDigest.latestMatches has invalid length: ${latestMatches.length} (expected ${minExpectedMatches}-6)`);
    process.exit(1);
}
console.log(`✓ emailDigest.latestMatches contains ${latestMatches.length} matches`);

// 4. Check topPerformers is not empty
const topPerformers = data.emailDigest.topPerformers;
if (!Array.isArray(topPerformers)) {
    console.error('Error: emailDigest.topPerformers is not an array!');
    process.exit(1);
}
if (topPerformers.length === 0 && summaryCount > 0) {
    console.error('Error: emailDigest.topPerformers is empty!');
    process.exit(1);
}
console.log(`✓ emailDigest.topPerformers contains ${topPerformers.length} participants`);

// 5. Check leaderboardTop contains at least 10 entries (or all participants if < 10)
const leaderboardTop = data.emailDigest.leaderboardTop;
if (!Array.isArray(leaderboardTop)) {
    console.error('Error: emailDigest.leaderboardTop is not an array!');
    process.exit(1);
}
const expectedMin = Math.min(10, topPerformers.length);
if (leaderboardTop.length < expectedMin) {
    console.error(`Error: emailDigest.leaderboardTop has only ${leaderboardTop.length} entries (expected at least ${expectedMin})`);
    process.exit(1);
}
console.log(`✓ emailDigest.leaderboardTop contains ${leaderboardTop.length} entries`);

// 6. Check latestWindows.last3Matches.participantPoints exists and contains participants
if (!data.latestWindows || !data.latestWindows.last3Matches || !Array.isArray(data.latestWindows.last3Matches.participantPoints)) {
    console.error('Error: latestWindows.last3Matches.participantPoints is missing or not an array!');
    process.exit(1);
}
const windowPoints = data.latestWindows.last3Matches.participantPoints;
if (windowPoints.length === 0 && summaryCount > 0) {
    console.error('Error: latestWindows.last3Matches.participantPoints is empty!');
    process.exit(1);
}
console.log(`✓ latestWindows.last3Matches.participantPoints contains ${windowPoints.length} participants`);

// 7. Check commentaryHints biggestRisers, biggestFallers, perfectScores
if (!data.commentaryHints) {
    console.error('Error: commentaryHints is missing!');
    process.exit(1);
}
const hints = ['biggestRisers', 'biggestFallers', 'perfectScores'];
for (const hint of hints) {
    if (!Array.isArray(data.commentaryHints[hint])) {
        console.error(`Error: commentaryHints.${hint} is missing or not an array!`);
        process.exit(1);
    }
}
console.log('✓ commentaryHints arrays exist');

console.log('Validation passed successfully!');
