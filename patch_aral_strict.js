const fs = require('fs');
let code = fs.readFileSync('Code.js', 'utf8');

const replaceTarget = `      const isAralEnd = subject.match(/\\bARAL\\b/i) && !subject.toLowerCase().includes('panlipunan');

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
      }`;

const replacement = `      const isAralEnd = subject.match(/\\bARAL\\b/i) && !subject.toLowerCase().includes('panlipunan');

      if (isAralEnd) {
        // ARAL is strictly for 3-4 PM only
        candidateSlots = [
          { in: '3:00 PM',  out: '4:00 PM',  s: 900, e: 960 }
        ];
      }`;

code = code.replace(replaceTarget, replacement);

fs.writeFileSync('Code.js', code);
console.log('ARAL 3-4pm strict logic applied.');
