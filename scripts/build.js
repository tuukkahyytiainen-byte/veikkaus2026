const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');

console.log('Starting cross-platform build...');

// Create public directory if it doesn't exist
const publicDir = path.join(__dirname, '..', 'public');
if (!fs.existsSync(publicDir)) {
    fs.mkdirSync(publicDir, { recursive: true });
    console.log('Created public/ directory');
}

// Files to copy to public/
const filesToCopy = [
    'index.html',
    'MM2026_pistelaskenta.xlsx',
    'tulokset.txt',
    'botmanen.jpg'
];

filesToCopy.forEach(file => {
    const src = path.join(__dirname, '..', file);
    const dest = path.join(publicDir, file);
    if (fs.existsSync(src)) {
        fs.copyFileSync(src, dest);
        console.log(`Copied ${file} to public/`);
    } else {
        console.warn(`Warning: Source file ${file} not found!`);
    }
});

// Run generate-agent-analysis.js
const generatorScript = path.join(__dirname, 'generate-agent-analysis.js');
console.log('Running generate-agent-analysis.js...');
try {
    execSync(`node "${generatorScript}"`, { stdio: 'inherit' });
    console.log('Agent analysis generated successfully.');
} catch (error) {
    console.error('Error running generator script:', error);
    process.exit(1);
}

// Run validation script if it exists
const validatorScript = path.join(__dirname, 'validate-agent-analysis.js');
if (fs.existsSync(validatorScript)) {
    console.log('Running validation...');
    try {
        execSync(`node "${validatorScript}"`, { stdio: 'inherit' });
        console.log('Validation successful.');
    } catch (error) {
        console.error('Validation failed:', error);
        process.exit(1);
    }
}

console.log('Build completed successfully!');
