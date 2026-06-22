const fs = require('fs');
let code = fs.readFileSync('Code.js', 'utf8');

// 1. Strictly follow weekly class hours: let hoursLeft = parseFloat(d[3]) || 0;
code = code.replace(/let hoursLeft = Math\.round\(parseFloat\(d\[3\]\) \|\| 0\);/, 'let hoursLeft = parseFloat(d[3]) || 0;');

fs.writeFileSync('Code.js', code);
console.log('Patch applied.');
