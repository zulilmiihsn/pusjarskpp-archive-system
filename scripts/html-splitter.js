const fs = require('fs');
const path = require('path');

const root = path.resolve(__dirname, '..');

// 1. Split Styles.html
const stylesContent = fs.readFileSync(path.join(root, 'Styles.html'), 'utf8');
const lines = stylesContent.split('\n');

const chunks = [];
let currentChunk = [];
let braceCount = 0;

// Remove <style> and </style>
const cssLines = lines.slice(1, -2);

const targetChunkSize = Math.ceil(cssLines.length / 5);

for (let i = 0; i < cssLines.length; i++) {
  const line = cssLines[i];
  currentChunk.push(line);
  
  // Very rough brace counting
  braceCount += (line.match(/\{/g) || []).length;
  braceCount -= (line.match(/\}/g) || []).length;

  if (currentChunk.length >= targetChunkSize && braceCount === 0) {
    chunks.push(currentChunk.join('\n'));
    currentChunk = [];
  }
}
if (currentChunk.length > 0) {
  chunks.push(currentChunk.join('\n'));
}

const styleNames = ['StylesBase', 'StylesLayout', 'StylesComponents', 'StylesForms', 'StylesTables'];
for (let i = 0; i < chunks.length; i++) {
  const name = styleNames[i] || `StylesExt${i}`;
  fs.writeFileSync(path.join(root, `${name}.html`), `<style>\n${chunks[i]}\n</style>`);
  console.log(`Created ${name}.html`);
}

// 2. Split ClientSettings.html
const settingsContent = fs.readFileSync(path.join(root, 'ClientSettings.html'), 'utf8');
const settingsChunks = settingsContent.split('function settings');
fs.writeFileSync(path.join(root, 'ClientSettingsGeneral.html'), `<script>\nfunction settings${settingsChunks[1]}`);
fs.writeFileSync(path.join(root, 'ClientSettingsSync.html'), `<script>\nfunction settings${settingsChunks[2]}`);
fs.writeFileSync(path.join(root, 'ClientSettingsMapping.html'), `<script>\nfunction settings${settingsChunks[3]}`);
fs.writeFileSync(path.join(root, 'ClientSettingsSystem.html'), `<script>\n${settingsChunks[0].replace('<script>\n', '')}\nfunction settings${settingsChunks[4] || ''}`);
console.log('Split ClientSettings');

// 3. Split ClientProcess.html
const processContent = fs.readFileSync(path.join(root, 'ClientProcess.html'), 'utf8');
const processChunks = processContent.split('function process');
fs.writeFileSync(path.join(root, 'ClientProcessUpload.html'), `<script>\nfunction process${processChunks[1]}`);
fs.writeFileSync(path.join(root, 'ClientProcessMetadata.html'), `<script>\nfunction process${processChunks[2]}`);
fs.writeFileSync(path.join(root, 'ClientProcessSubmit.html'), `<script>\n${processChunks[0].replace('<script>\n', '')}\nfunction process${processChunks[3] || ''}`);
console.log('Split ClientProcess');
