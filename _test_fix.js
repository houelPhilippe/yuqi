// Test the unescape fix for getCellParas
const cases = [
  '------ \u2713 TransitOperation',
  '------ ------ \u2713 MRN',
  '------ ------ ------ \u2713 code',
  '- text',      // single dash + space (should stay escaped by turndown)
  '-- text',     // two dashes + space (unescaped by fix, but not a list marker)
  '\\------ \u2713 TransitOperation',  // already escaped (from CC042Cb.md)
];

// Simulate turndown escape function (/^-/mg → '\-')
function simulateTurndownEscape(text) {
  return text.replace(/^-/mg, '\\-');
}

// Our fix: unescape backslash before 2+ consecutive dashes
function applyFix(md) {
  return md.replace(/^\\(-{2,})/mg, '$1');
}

cases.forEach(c => {
  const afterTurndown = simulateTurndownEscape(c);
  const afterFix = applyFix(afterTurndown);
  console.log('Input   :', JSON.stringify(c));
  console.log('Turndown:', JSON.stringify(afterTurndown));
  console.log('Fixed   :', JSON.stringify(afterFix));
  console.log('');
});
