const fs = require('fs');
let code = fs.readFileSync('Code.js', 'utf8');

// We need to add Phil Gov detection.
// Right before: let candidateSlots = grade >= 11 ? SHS_SLOTS : STANDARD_SLOTS;

const targetPhil = `      let candidateSlots = grade >= 11 ? SHS_SLOTS : STANDARD_SLOTS;`;

const replacementPhil = `      let candidateSlots = grade >= 11 ? SHS_SLOTS : STANDARD_SLOTS;

      const isPhilGov = subject.toLowerCase().includes('phil gov') || subject.toLowerCase().includes('philippine politics');
      if (isPhilGov) {
        // Force strict 90-minute (1.5h) blocks for Phil Gov
        candidateSlots = [
          { in: '7:30 AM',  out: '9:00 AM',  s: 450, e: 540 },
          { in: '9:00 AM',  out: '10:30 AM', s: 540, e: 630 },
          { in: '10:45 AM', out: '12:15 PM', s: 645, e: 735 }, // Crosses lunch slightly if lunch starts at 11:45, wait lunch is 11:45-1:00.
          // Let's adjust to be safe around lunch (11:45-1:00)
          // 7:30 to 9:00
          // 9:00 to 10:30
          // 10:30 to 11:45 is only 1.25 hours (75 mins), so can't fit a 90 min block before lunch.
          // After lunch: 1:00 to 2:30 PM, 2:30 to 4:00 PM
          { in: '1:00 PM',  out: '2:30 PM',  s: 780, e: 870 },
          { in: '2:30 PM',  out: '4:00 PM',  s: 870, e: 960 }
        ];
        // Ensure that these slots are only up to 10:30 AM to respect the 11:45 lunch, so we drop the 10:45 slot
        // since 10:45 to 12:15 overlaps with lunch.
        candidateSlots = [
          { in: '7:30 AM',  out: '9:00 AM',  s: 450, e: 540 },
          { in: '9:00 AM',  out: '10:30 AM', s: 540, e: 630 },
          { in: '1:00 PM',  out: '2:30 PM',  s: 780, e: 870 },
          { in: '2:30 PM',  out: '4:00 PM',  s: 870, e: 960 }
        ];
      }`;

code = code.replace(targetPhil, replacementPhil);

fs.writeFileSync('Code.js', code);
console.log('Phil Gov logic added.');
