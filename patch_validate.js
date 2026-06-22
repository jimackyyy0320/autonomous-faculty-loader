const fs = require('fs');
let code = fs.readFileSync('Code.js', 'utf8');

// Ensure hoursLeft deduction uses floating point properly and stops at 0
// Check the "bookSlot" logic
let match = code.match(/const bookSlot = \([\s\S]*?\};/);
if (match) {
    console.log(match[0]);
}
