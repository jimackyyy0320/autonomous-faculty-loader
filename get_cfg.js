const fs = require('fs');
let code = fs.readFileSync('Code.js', 'utf8');
console.log(code.match(/const CFG = \{[\s\S]*?\};/)[0]);
