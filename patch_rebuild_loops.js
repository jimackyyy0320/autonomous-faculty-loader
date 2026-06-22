const fs = require('fs');
let code = fs.readFileSync('Code.js', 'utf8');

const regex = /for \(let day of prefDays\) \{[\s\S]*?if \(hoursLeft > 0\) \{/m;

const replacement = `for (let day of prefDays) {
        if (hoursLeft <= 0) break;
        for (let slot of candidateSlots) {
          if (daysUsedForSubject.has(day) && !isHomeroom) continue;
          if (isFree(day, slot, true)) { bookSlot(day, slot); break; }
        }
      }

      if (hoursLeft > 0 && !isHomeroom) {
        for (let day of [1,2,3,4,5]) {
          if (hoursLeft <= 0) break;
          if (slotsAcquired.some(s => s.day === day)) continue;
          for (let slot of candidateSlots) {
            if (daysUsedForSubject.has(day)) continue;
            if (isFree(day, slot, true)) { bookSlot(day, slot); break; }
          }
        }
      }

      // Fallback: Ignore preferred hours check
      if (hoursLeft > 0 && !isHomeroom) {
        for (let day of prefDays) {
          if (hoursLeft <= 0) break;
          for (let slot of candidateSlots) {
            if (daysUsedForSubject.has(day)) continue;
            if (isFree(day, slot, false)) {
              bookSlot(day, slot);
              const warnStr = \`⚠️ Exceeds \${CFG.DAILY_PREFERRED_HOURS}h limit\`;
              unmappedLog.push(\`[\${term}] Soft Warning: \${teacher} exceeded \${CFG.DAILY_PREFERRED_HOURS}h preferred limit to accommodate \${subject} (\${section}).\`);
              if (!rowWarnings.includes(warnStr)) rowWarnings.push(warnStr);
              break;
            }
          }
        }
      }

      if (hoursLeft > 0 && !isHomeroom) {
        for (let day of [1,2,3,4,5]) {
          if (hoursLeft <= 0) break;
          for (let slot of candidateSlots) {
            if (hoursLeft <= 0) break;
            if (daysUsedForSubject.has(day)) continue;
            if (isFree(day, slot, false)) {
              bookSlot(day, slot);
              const warnStr = \`⚠️ Exceeds \${CFG.DAILY_PREFERRED_HOURS}h limit\`;
              unmappedLog.push(\`[\${term}] Soft Warning: \${teacher} exceeded \${CFG.DAILY_PREFERRED_HOURS}h preferred limit to accommodate \${subject} (\${section}).\`);
              if (!rowWarnings.includes(warnStr)) rowWarnings.push(warnStr);
              break;
            }
          }
        }
      }

      if (hoursLeft > 0) {`;

code = code.replace(regex, replacement);

// We need to fix bookSlot to truncate if hoursLeft < duration
const bookSlotRegex = /const bookSlot = \([\s\S]*?\};/;
const newBookSlot = `const bookSlot = (day, slot) => {
        let dur = (slot.e - slot.s) / 60;
        let actualE = slot.e;
        let actualOut = slot.out;
        if (hoursLeft < dur) {
            // Truncate the slot
            actualE = slot.s + Math.round(hoursLeft * 60);
            actualOut = formatMinsToTime(actualE);
            dur = hoursLeft;
        }

        tBooked[teacher][day].push({s: slot.s, e: actualE});
        tBookedMins[teacher][day] += (actualE - slot.s);
        cBooked[section][day].push({s: slot.s, e: actualE});
        slotsAcquired.push({ day, in: slot.in, out: actualOut, s: slot.s, e: actualE });
        hoursLeft -= dur;
        daysUsedForSubject.add(day);
      };`;

code = code.replace(bookSlotRegex, newBookSlot);

// Also formatMinsToTime needs to be available in this scope, but it's a global function in Code.js so it's fine.

fs.writeFileSync('Code.js', code);
console.log('BookSlot truncated, loops rebuilt.');
