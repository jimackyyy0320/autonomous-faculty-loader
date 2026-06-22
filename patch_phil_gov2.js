const fs = require('fs');
let code = fs.readFileSync('Code.js', 'utf8');

// There is one more edge case with Phil Gov. If we have to split 3 hours, we should only use prefDays.
// Is there anything we need to do to the prefDays allocation?
// It defaults to 3 hours = [1,3,5]. So that's perfectly 3 days of 90 mins each. Wait, 3 hours total?
// Phil Gov is usually a 3 unit subject. If 3 hours / 1.5h = 2 days!
// So if hoursLeft === 3, prefDays should just be [2,4] or [1,3] since it only needs 2 slots.
// Let's modify the prefDays assignment logic to handle Phil Gov specifically.

const targetPrefDays = `        if (hoursLeft === 4) prefDays = [1,2,4,5];
        if (hoursLeft === 3) prefDays = [1,3,5];
        if (hoursLeft === 2) prefDays = [2,4];
        if (hoursLeft === 1) prefDays = [3];
        prefDays = prefDays.sort(() => Math.random() - 0.5);`;

const replacementPrefDays = `        if (isPhilGov && hoursLeft === 3) {
          prefDays = [2, 4]; // 2 days x 1.5 hours = 3 hours
        } else {
          if (hoursLeft === 4) prefDays = [1,2,4,5];
          if (hoursLeft === 3) prefDays = [1,3,5];
          if (hoursLeft === 2) prefDays = [2,4];
          if (hoursLeft === 1) prefDays = [3];
        }
        prefDays = prefDays.sort(() => Math.random() - 0.5);`;

code = code.replace(targetPrefDays, replacementPrefDays);

fs.writeFileSync('Code.js', code);
console.log('Phil Gov prefDays logic added.');
