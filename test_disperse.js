// The logic is already present, but the prompt emphasizes "disperse it... instead of congesting loads of teacher on a straight morning or straight afternoon".
// We applied a 3.5 limit. Let's make it slightly stricter to 3 hours to force better dispersion, but only when checking preferred.
const fs = require('fs');
let code = fs.readFileSync('Code.js', 'utf8');

code = code.replace(/3\.5 \* 60/g, '3 * 60'); // 3 hours limit per half day

fs.writeFileSync('Code.js', code);
console.log('Stricter dispersion applied.');
