const fs = require('fs');
let code = fs.readFileSync('Code.js', 'utf8');

// The match was found, so the replace in test_slots.js failed. Let's force it properly.
const regex = /\/\/ Shuffle candidateSlots slightly to allow for regeneration randomness\n\s*candidateSlots = \[\.\.\.candidateSlots\]\.sort\(\(\) => Math\.random\(\) - 0\.5\);/;

const replacement = `// Keep candidateSlots sequential to ensure students' classes are full from 7:30 AM onwards.
      // We only randomize days to distribute load, but times should be filled sequentially.`;

code = code.replace(regex, replacement);

fs.writeFileSync('Code.js', code);
console.log('Shuffle actually removed.');
