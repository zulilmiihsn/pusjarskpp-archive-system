const fs = require('fs');
const path = require('path');

const logoPath = path.join(__dirname, '../assets/logo PUSJARSKPP-01.png');
const gedungPath = path.join(__dirname, '../assets/gedung utama.png');
const ornamenPath = path.join(__dirname, '../assets/ornamen.png');
const logoSosmedPath = path.join(__dirname, '../assets/LOGO SOSMED 2025-03.png');
const lightOutputPath = path.join(__dirname, '../ClientAssets.html');
const heavyOutputPath = path.join(__dirname, '../ClientAssetsHeavy.html');

const dataUri = (p) => 'data:image/png;base64,' + fs.readFileSync(p).toString('base64');

console.log('Membaca aset gambar...');
const logo = dataUri(logoPath);
const logoSosmed = dataUri(logoSosmedPath);
const gedungUtama = dataUri(gedungPath);
const ornamen = dataUri(ornamenPath);

// Bundle KRITIS (logo + watermark sosmed): di-inline ke Index.html, dipakai
// pada setiap halaman untuk branding. Kecil (~130KB).
const lightContent = `<script>
  const ASSETS = {
    logo: '${logo}',
    logoSosmed: '${logoSosmed}'
  };
</script>
`;

// Bundle BERAT (background dekoratif): TIDAK di-inline. Hanya dimuat lazy lewat
// endpoint getDecorativeAssets saat toggle background aktif. Format JSON murni
// supaya server bisa langsung JSON.parse isinya. ~2.5MB → keluar dari payload awal.
const heavyContent = JSON.stringify({ gedungUtama: gedungUtama, ornamen: ornamen });

console.log('Menulis ClientAssets.html (kritis) & ClientAssetsHeavy.html (lazy)...');
fs.writeFileSync(lightOutputPath, lightContent, 'utf8');
fs.writeFileSync(heavyOutputPath, heavyContent, 'utf8');
console.log('Aset berhasil dibuat (kritis + lazy terpisah)!');
