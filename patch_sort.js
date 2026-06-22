const fs = require('fs');
let code = fs.readFileSync('Code.js', 'utf8');

// If we sort by section first, we guarantee that we plot a whole section's schedule before moving to the next section.
// This greatly increases the chance that a section's schedule is contiguous because it claims all the earliest slots for itself before competing with other sections.
const targetSort = `    // Prioritize Homerooms, then heaviest hours
    termDemands.sort((a, b) => {
      const isAHomeroom = a[2].toString().toLowerCase().includes('homeroom') ? 1 : 0;
      const isBHomeroom = b[2].toString().toLowerCase().includes('homeroom') ? 1 : 0;
      if (isAHomeroom !== isBHomeroom) return isBHomeroom - isAHomeroom;
      return (parseFloat(b[3]) || 0) - (parseFloat(a[3]) || 0);
    });`;

const replacementSort = `    // Prioritize Homerooms, then group by Section, then by heaviest hours to pack student schedules continuously
    termDemands.sort((a, b) => {
      const isAHomeroom = a[2].toString().toLowerCase().includes('homeroom') ? 1 : 0;
      const isBHomeroom = b[2].toString().toLowerCase().includes('homeroom') ? 1 : 0;
      if (isAHomeroom !== isBHomeroom) return isBHomeroom - isAHomeroom;

      const sectionA = a[1].toString().trim();
      const sectionB = b[1].toString().trim();
      if (sectionA !== sectionB) return sectionA.localeCompare(sectionB);

      return (parseFloat(b[3]) || 0) - (parseFloat(a[3]) || 0);
    });`;

code = code.replace(targetSort, replacementSort);

fs.writeFileSync('Code.js', code);
console.log('Sort updated to group by section.');
