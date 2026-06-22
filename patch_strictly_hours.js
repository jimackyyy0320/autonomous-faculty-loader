const fs = require('fs');
let code = fs.readFileSync('Code.js', 'utf8');

// The original isFree takes slot. We shouldn't change isFree too much, but we need to adjust bookSlot and the slot object passed to it.
// We can modify the loop where it checks isFree and calls bookSlot.
// Wait, candidateSlots is iterated. Let's intercept the slot before passing to isFree and bookSlot.

const replaceTarget = `
      const isFree = (day, slot, checkPreferred = true) => {
        const tConflict = tBooked[teacher][day].some(b => slot.s < b.e && slot.e > b.s);
        const cConflict = cBooked[section][day].some(b => slot.s < b.e && slot.e > b.s);
        if (tConflict || cConflict) return false;
`;

const replacement = `
      const isFree = (day, slot, checkPreferred = true) => {
        const tConflict = tBooked[teacher][day].some(b => slot.s < b.e && slot.e > b.s);
        const cConflict = cBooked[section][day].some(b => slot.s < b.e && slot.e > b.s);
        if (tConflict || cConflict) return false;

        // Anti-Congestion: prevent straight mornings or afternoons without a break
        // Check how many hours this teacher already has in the morning (before 12:00 PM/720 mins)
        // or afternoon (after 12:00 PM). Soft limit to 3.5 hours straight per half-day.
        if (checkPreferred && !isHomeroom) {
            const isMorning = slot.s < 720;
            let halfDayMins = 0;
            tBooked[teacher][day].forEach(b => {
                if (isMorning && b.s < 720) halfDayMins += (b.e - b.s);
                else if (!isMorning && b.s >= 720) halfDayMins += (b.e - b.s);
            });
            if (halfDayMins + (slot.e - slot.s) > 3.5 * 60) return false;
        }
`;

code = code.replace(replaceTarget, replacement);

const replaceTargetLoop = `
      for (let day of prefDays) {
        if (hoursLeft <= 0) break;
        for (let slot of candidateSlots) {
          if (isFree(day, slot, true)) { bookSlot(day, slot); break; }
        }
      }
`;

// To prevent same subject same section on same day, we need to track if we already plotted this subject-section on this day.
// We can use a Set: \`let daysUsed = new Set();\` per subject-section.
// Wait, the loops are inside the \`termDemands.forEach(d => { ... })\`, so we can just declare \`let daysUsedForSubject = new Set();\` inside the \`forEach\`.

fs.writeFileSync('Code.js', code);
console.log('Anti-congestion applied in isFree.');
