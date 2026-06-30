const s = '\\------ ✓ TransitOperation';
const r = /^\\(-{2,})/gm;
console.log("Original: " + s);
console.log("Replaced: " + s.replace(r, '$1'));
