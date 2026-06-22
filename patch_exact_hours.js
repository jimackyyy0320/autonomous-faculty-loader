const fs = require('fs');
let code = fs.readFileSync('Code.js', 'utf8');

// We need to truncate the slot duration to exact hours left.
const replaceTarget = `
        let isFreeOnAllDays = true;
`;

// Wait, the actual loops using bookSlot:
const targetLoops = `
      for (let day of prefDays) {
        if (hoursLeft <= 0) break;
        for (let slot of candidateSlots) {
          if (daysUsedForSubject.has(day) && !isHomeroom) continue; // Prevent same day duplication
          if (isFree(day, slot, true)) { bookSlot(day, slot); break; }
        }
      }

      if (hoursLeft > 0 && !isHomeroom) {
        for (let day of [1,2,3,4,5]) {
          if (hoursLeft <= 0) break;
          if (slotsAcquired.some(s => s.day === day)) continue;
          for (let slot of candidateSlots) {
            if (daysUsedForSubject.has(day)) continue; // Prevent same day duplication
            if (isFree(day, slot, true)) { bookSlot(day, slot); break; }
          }
        }
      }

      // Fallback: Ignore preferred hours check
      if (hoursLeft > 0 && !isHomeroom) {
        for (let day of prefDays) {
          if (hoursLeft <= 0) break;
          for (let slot of candidateSlots) {
            if (daysUsedForSubject.has(day)) continue; // Prevent same day duplication
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
            if (daysUsedForSubject.has(day)) continue; // Prevent same day duplication
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
`;

// Let's replace the loops directly in Code.js using a script to rebuild the loops block.
