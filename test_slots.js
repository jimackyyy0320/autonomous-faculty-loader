const fs = require('fs');
let code = fs.readFileSync('Code.js', 'utf8');

// The prompt mentions:
// "Student classess should be full from 7:30 to 3 or 4pm. Students should be inside their classrooms during class hours."
// To achieve this, let's remove the shuffle on candidateSlots for standard subjects and just keep them sequential so it plots earlier slots first.

const targetCandidateShuffle = `      // Shuffle candidateSlots slightly to allow for regeneration randomness
      candidateSlots = [...candidateSlots].sort(() => Math.random() - 0.5);`;

code = code.replace(targetCandidateShuffle, `      // Keep candidateSlots sequential to ensure students' classes are full from 7:30 AM onwards, avoiding gaps in their schedule.
      // We only randomize days to distribute load, but times should be filled sequentially.
      // Note: We don't shuffle candidateSlots here anymore except for Phil Gov specific slots which will be handled.
      // candidateSlots = [...candidateSlots].sort(() => Math.random() - 0.5);`);

fs.writeFileSync('Code.js', code);
console.log('Candidate slots sequential.');
