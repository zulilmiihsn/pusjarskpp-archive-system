const fs = require('fs');
const path = require('path');

const dir = path.resolve(__dirname, '..');
const files = fs.readdirSync(dir);

let syncedCount = 0;
for (const file of files) {
  if (file.endsWith('.gs')) {
    const gsPath = path.join(dir, file);
    const jsPath = path.join(dir, file.replace(/\.gs$/, '.js'));
    const content = fs.readFileSync(gsPath, 'utf8');
    fs.writeFileSync(jsPath, content, 'utf8');
    syncedCount++;
  }
}
console.log(`Synced ${syncedCount} .gs files to .js files.`);
