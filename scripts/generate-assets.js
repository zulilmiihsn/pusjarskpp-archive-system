const fs = require('fs');
const path = require('path');

const logoPath = path.join(__dirname, '../assets/logo PUSJARSKPP-01.png');
const gedungPath = path.join(__dirname, '../assets/gedung utama.png');
const ornamenPath = path.join(__dirname, '../assets/ornamen.png');
const logoSosmedPath = path.join(__dirname, '../assets/LOGO SOSMED 2025-03.png');
const outputPath = path.join(__dirname, '../ClientAssets.html');

console.log('Membaca file logo dan gedung...');
const logoBase64 = fs.readFileSync(logoPath).toString('base64');
const gedungBase64 = fs.readFileSync(gedungPath).toString('base64');
const ornamenBase64 = fs.readFileSync(ornamenPath).toString('base64');
const logoSosmedBase64 = fs.readFileSync(logoSosmedPath).toString('base64');

const htmlContent = `<script>
  const ASSETS = {
    logo: 'data:image/png;base64,${logoBase64}',
    gedungUtama: 'data:image/png;base64,${gedungBase64}',
    ornamen: 'data:image/png;base64,${ornamenBase64}',
    logoSosmed: 'data:image/png;base64,${logoSosmedBase64}'
  };
</script>
`;

console.log('Menulis file ClientAssets.html...');
fs.writeFileSync(outputPath, htmlContent, 'utf8');
console.log('File ClientAssets.html berhasil dibuat!');
