const text = 'Nomor : 267/P.3/PDP.07.1';
const r1 = text.match(/(?:No(?:mor)?)\s*[:.\s]+\s*([A-Z0-9][A-Z0-9.\/\-]+\/[A-Z0-9.\/\-]+)/i);
console.log('r1:', r1);

const r2 = text.match(/(?:No(?:mor)?)\s*[:.\-\s_]+\s*([A-Z0-9][A-Z0-9.\-\s_]+?[\-\s_][12]\d{3})\b/i);
console.log('r2:', r2);
