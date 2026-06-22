const fs = require('fs');
let code = fs.readFileSync('Code.js', 'utf8');

// The new requirement: "ARAL" subject should be at the end of their class schedule.
// "ARAL" and "Araling Panlipunan" are different. We need an exact match for the word "ARAL".

const replaceTarget = `      if (isHomeroom) {`;

const replacement = `      const isAralEnd = subject.match(/\\bARAL\\b/i) && !subject.toLowerCase().includes('panlipunan');

      if (isAralEnd) {
        if (grade >= 11) {
          candidateSlots = [
            { in: '2:30 PM',  out: '4:00 PM',  s: 870, e: 960 }
          ];
        } else {
          candidateSlots = [
            { in: '2:00 PM',  out: '3:00 PM',  s: 840, e: 900 },
            { in: '3:00 PM',  out: '4:00 PM',  s: 900, e: 960 }
          ];
        }
      }

      if (isHomeroom) {`;

code = code.replace(replaceTarget, replacement);

fs.writeFileSync('Code.js', code);
console.log('ARAL end-of-day logic added.');
