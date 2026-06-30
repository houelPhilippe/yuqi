const marked = require('./frontend/marked.min.js');
marked.setOptions({ breaks: true, gfm: true });

// Cas réels du fichier CC042C.md
const cases = [
  // Contenu exact des cellules de la 1ère colonne
  '------ \u2713 TransitOperation',
  '------ ------ \u2713 MRN',
  '------ ------ ------ \u2713 code',
  // Après turndown escape (ce qu'on trouve dans CC042Cb.md)
  '\\------ \u2713 TransitOperation',
  '\\------ ------ \u2713 MRN',
  '\\------ ------ ------ \u2713 code',
  // Variantes avec espace après le premier tiret
  '- ----- \u2713 TransitOperation',
];

console.log('=== Test des tirets sans espaces ===\n');
cases.forEach(c => {
  const html = marked.parse(c);
  const isHr       = html.includes('<hr');
  const isList     = html.includes('<ul>') || html.includes('<li>');
  const hasBr      = /<br\s*\/?>/.test(html);
  const isMultiTag = (html.match(/<[a-z]/g) || []).length > 2;
  console.log('Input :', JSON.stringify(c));
  console.log('HTML  :', html.replace(/\n/g,'↵').trim().substring(0, 100));
  console.log('→ HR:', isHr, '| List:', isList, '| BR:', hasBr, '| Multi-blocs:', isMultiTag);
  console.log('');
});
