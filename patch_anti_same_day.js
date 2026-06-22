const fs = require('fs');
let code = fs.readFileSync('Code.js', 'utf8');

const replaceTarget = `
      let slotsAcquired = [];
      let rowWarnings = [];
`;

const replacement = `
      let slotsAcquired = [];
      let rowWarnings = [];
      let daysUsedForSubject = new Set();
`;

code = code.replace(replaceTarget, replacement);

const replaceTargetBookSlot = `
      const bookSlot = (day, slot) => {
        tBooked[teacher][day].push({s: slot.s, e: slot.e});
        tBookedMins[teacher][day] += (slot.e - slot.s);
        cBooked[section][day].push({s: slot.s, e: slot.e});
        slotsAcquired.push({ day, in: slot.in, out: slot.out, s: slot.s });
        hoursLeft -= ((slot.e - slot.s) / 60);
      };
`;

const replacementBookSlot = `
      const bookSlot = (day, slot) => {
        tBooked[teacher][day].push({s: slot.s, e: slot.e});
        tBookedMins[teacher][day] += (slot.e - slot.s);
        cBooked[section][day].push({s: slot.s, e: slot.e});
        slotsAcquired.push({ day, in: slot.in, out: slot.out, s: slot.s });
        hoursLeft -= ((slot.e - slot.s) / 60);
        daysUsedForSubject.add(day);
      };
`;

code = code.replace(replaceTargetBookSlot, replacementBookSlot);

fs.writeFileSync('Code.js', code);
console.log('daysUsedForSubject added.');
