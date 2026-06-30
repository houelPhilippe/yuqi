const marked = require('./frontend/marked.min.js');
marked.setOptions({ breaks: true, gfm: true });

// Simuler une ligne de tableau pipe (CC042C.md)
const pipeLine = '| ------ ✓ TransitOperation | Obligatoire |';
const pipeTable = `| Élément | Statut |
| --- | --- |
| ✓ CC042C | Obligatoire |
| ------ ✓ TransitOperation | Obligatoire |
| ------ ------ ✓ MRN | Obligatoire |
| ------ ------ ------ ✓ code | Obligatoire |`;

const html = marked.parse(pipeTable);
console.log('=== Rendu du tableau pipe ===');
console.log(html);
console.log('');

// Isoler chaque cellule
const parser = new DOMParser ? null : null;
// Montrer uniquement les cellules de la 1ère colonne
const cellRe = /<td[^>]*>([\s\S]*?)<\/td>/g;
let m;
let i = 0;
while ((m = cellRe.exec(html)) !== null) {
  if (i % 2 === 0) {
    console.log(`Cell ${i/2} innerHTML: ${JSON.stringify(m[1])}`);
  }
  i++;
}
