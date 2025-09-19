const { execSync } = require('child_process');
const fs = require('fs');
const path = require('path');

console.log('Building Thai Post Check application...');

try {
  // Check if electron-builder is installed
  execSync('electron-builder --version', { stdio: 'ignore' });
  console.log('electron-builder found');
} catch (error) {
  console.error('electron-builder not found. Please install it first:');
  console.error('bun add --dev electron-builder');
  process.exit(1);
}

// Create dist directory if it doesn't exist
const distDir = path.join(__dirname, 'dist');
if (!fs.existsSync(distDir)) {
  fs.mkdirSync(distDir);
  console.log('Created dist directory');
}

// Run electron-builder
console.log('Running electron-builder...');
execSync('electron-builder --win', { stdio: 'inherit' });

console.log('Build completed successfully!');
console.log('The executable can be found in the dist directory.');