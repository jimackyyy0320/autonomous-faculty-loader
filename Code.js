/**
 * ============================================================
 * AUTONOMOUS FACULTY LOADING SYSTEM — MASTER V4
 * Google Apps Script — Single File Deployment
 * ============================================================
 * NEW IN V4: Section (Cohort) Dashboard & Chronological Sorting
 * ============================================================
 */

// ══════════════════════════════════════════════════════════════
//  SECTION 1: CONFIGURATION & PALETTE
// ══════════════════════════════════════════════════════════════

const CFG = {
  SECTION_ENROLL: 'Section Enrollment',
  SUBJECT_LOAD:   'Subject Loading',
  TEACHER_ENROLL: 'Teacher Enrollment',
  TEACHER_ASSIGN: 'Teacher Assignment',
  TERMS:          ['Term 1', 'Term 2', 'Term 3'],
  DASHBOARD:      'Teacher Dashboard',
  SECTION_DASH:   'Section Dashboard', // NEW
  REPORT:         'Schedule Alignment',
  
  LUNCH_START:       705,  // 11:45 AM (mins since midnight)
  LUNCH_END:         780,  // 1:00 PM
  WEEKLY_WARN_HOURS: 28,   
  WEEKLY_HARD_HOURS: 30,   
  DAILY_PREFERRED_HOURS: 5.5,
  DAILY_WARN_HOURS:  5.5, 
  DAILY_HARD_HOURS:  6,   
  SLOT_STEP:         30,   
  SCHOOL_START:      420,  // 7:00 AM
  SCHOOL_END:        990,  // 4:30 PM
};

const C = {
  navyDark: '#0d1b3e', navy:      '#1a3a6e', navyLight: '#e8edf7',
  teal:     '#00796b', tealLight: '#e0f2f1',
  error:    '#b71c1c', errorBg:   '#ffebee',
  warn:     '#e65100', warnBg:    '#fff3e0',
  info:     '#1565c0', infoBg:    '#e3f2fd',
  ok:       '#2e7d32', okBg:      '#e8f5e9',
  white:    '#ffffff', rowAlt:    '#f8f9fa',
  border:   '#cfd8dc', muted:     '#78909c', body: '#212121',
};

// ══════════════════════════════════════════════════════════════
//  SECTION 2: TRIGGERS (The Autonomy Engine)
// ══════════════════════════════════════════════════════════════

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📋 Faculty Tools')
    .addItem('1️⃣ Build: Phase 1 Data Tabs', 'buildPhase1Tabs')
    .addItem('2️⃣ Build: Phase 2 Teacher Tabs', 'buildPhase2TeacherTabs')
    .addItem('3️⃣ Build: Phase 3 Term Tabs', 'buildPhase3TermTabs')
    .addItem('4️⃣ Build: Phase 4 Teacher Dash', 'buildPhase4Dashboard')
    .addItem('5️⃣ Build: Phase 5 Conflict Report', 'buildPhase5ConflictReport')
    .addItem('6️⃣ Build: Phase 6 Section Dash', 'buildPhase6SectionDashboard') // NEW
    .addSeparator()
    .addItem('🚀 RUN AUTO-SCHEDULER', 'runAutoScheduler')
    .addItem('🔄 RECALCULATE SCHEDULE', 'runAutoScheduler')
    .addItem('🔍 RUN CONFLICT CHECKER', 'runConflictChecker')
    .addSeparator()
    .addItem('🖨️ Generate Schedule PDF', 'generateScheduleUI')
    .addToUi();
}

function onEdit(e) {
  if (!e || !e.range) return;
  
  const sheetName = e.range.getSheet().getName();
  const row = e.range.getRow();
  const col = e.range.getColumn();
  
  // 1. Foundation Sweeps & Syncs
  if (sheetName === CFG.SECTION_ENROLL || sheetName === CFG.SUBJECT_LOAD || sheetName === CFG.TEACHER_ENROLL) {
    cleanOrphanedData();
    if (sheetName === CFG.SECTION_ENROLL) updateSectionDropdowns();
    if (sheetName === CFG.SUBJECT_LOAD) updateTeacherDropdownOptions();
    if (sheetName === CFG.TEACHER_ENROLL) updateTeacherNameDropdowns(); 
  }

  // 2. Dashboards Live Filter (Both Teacher and Section)
  if (sheetName === CFG.SECTION_DASH && col === 2 && row >= 2 && row <= 4) {
    const sheet = e.range.getSheet();
    if (row === 3) { e.value === 'ALL WEEK' ? sheet.showColumns(5, 5) : sheet.hideColumns(5, 5); }
    updateSectionDashboardUI(sheet, SpreadsheetApp.getActiveSpreadsheet());
  }

  if (sheetName === CFG.DASHBOARD && (col === 2 || col === 12) && row >= 2 && row <= 4) {
    const sheet = e.range.getSheet();
    if (row === 3) {
      if (col === 2) { e.value === 'ALL WEEK' ? sheet.showColumns(5, 5) : sheet.hideColumns(5, 5); }
      if (col === 12) { e.value === 'ALL WEEK' ? sheet.showColumns(15, 5) : sheet.hideColumns(15, 5); }
    }
    updateDashboardUI(sheet, SpreadsheetApp.getActiveSpreadsheet(), col === 2 ? 1 : 2);
  }

  // 3. Conflict 1-Click "Implement" Auto-Fixer
  if (sheetName === CFG.REPORT && col === 9 && row > 2 && e.value === 'Implement') {
    applySuggestedFix(e);
  }

  // 4. Real-time Teacher Conflict Checker
  if (CFG.TERMS.includes(sheetName) && col === 11 && row > 2) {
    checkTeacherConflicts(e.range.getSheet(), e.value, row);
  }

  // 4.5 Term Tab Conflict Auto-Fixer
  if (CFG.TERMS.includes(sheetName) && col === 13 && row > 2 && e.value === 'Fix Conflict') {
    fixTermConflictTab(e, sheetName);
  }
}

// ══════════════════════════════════════════════════════════════
//  SECTION 3: AUTO-SCHEDULER (The Super-Advanced Brain)
// ══════════════════════════════════════════════════════════════

function runAutoScheduler() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const subSheet = ss.getSheetByName(CFG.SUBJECT_LOAD);
  const teachSheet = ss.getSheetByName(CFG.TEACHER_ENROLL);
  
  if (!subSheet || !teachSheet) {
    return ss.toast('Please ensure your Subject Loading and Teacher Enrollment tabs are set up first.', '👋 Just a quick note', 5);
  }

  ss.toast('Running Student-First Timetable Generator...', '🧠 Computing', 4);

  const getSafeData = (sheet, numCols) => {
    const lr = sheet.getLastRow();
    return lr < 3 ? [] : sheet.getRange(3, 1, lr - 2, numCols).getValues();
  };

  const demands = getSafeData(subSheet, 4).filter(r => r[0] && r[1] && r[2]); 
  const teachers = getSafeData(teachSheet, 3).filter(r => r[1]); // [Dept, Name, Specialization]

  // Clear Term tabs and rebuild data validations
  const teacherNames = teachers.map(t => t[1]);
  const teacherRule = SpreadsheetApp.newDataValidation().requireValueInList(teacherNames, true).build();

  CFG.TERMS.forEach(term => {
    const ts = ss.getSheetByName(term);
    if (ts) {
      const lr = ts.getLastRow();
      if (lr >= 3) {
        // Clear all data columns. New layout:
        // A: Section, B: Subject, C-G: MON-FRI, H: Time In, I: Time Out, J: Suggested, K: Assigned Teacher, L: Status
        ts.getRange(3, 1, lr - 2, 13).clearContent().clearDataValidations().clearFormat();
      }
    }
  });

  const STANDARD_SLOTS = [
    { in: '7:30 AM',  out: '8:30 AM',  s: 450, e: 510 },
    { in: '8:30 AM',  out: '9:30 AM',  s: 510, e: 570 },
    { in: '9:45 AM',  out: '10:45 AM', s: 585, e: 645 },
    { in: '10:45 AM', out: '11:45 AM', s: 645, e: 705 },
    { in: '1:00 PM',  out: '2:00 PM',  s: 780, e: 840 },
    { in: '2:00 PM',  out: '3:00 PM',  s: 840, e: 900 },
    { in: '3:00 PM',  out: '4:00 PM',  s: 900, e: 960 }
  ];

  const SHS_SLOTS = [
    { in: '7:30 AM',  out: '8:30 AM',  s: 450, e: 510 },
    { in: '8:30 AM',  out: '9:30 AM',  s: 510, e: 570 },
    { in: '9:45 AM',  out: '10:45 AM', s: 585, e: 645 },
    { in: '10:45 AM', out: '11:45 AM', s: 645, e: 705 },
    { in: '1:00 PM',  out: '2:30 PM',  s: 780, e: 870 },
    { in: '2:30 PM',  out: '4:00 PM',  s: 870, e: 960 }
  ];

  let unmappedLog = [];

  CFG.TERMS.forEach(term => {
    const termDemands = demands.filter(d => d[0] === term);
    if (!termDemands.length) return;

    // We no longer track tBooked during generation, only cBooked (Section Timetable)
    const cBooked = {}; 
    const outputRows = [];

    // Group by Section
    const sectionDemands = {};
    termDemands.forEach(d => {
      const sec = d[1].toString().trim();
      if (!sectionDemands[sec]) sectionDemands[sec] = [];
      sectionDemands[sec].push(d);
    });

    Object.keys(sectionDemands).forEach(section => {
      const sd = sectionDemands[section];
      
      // Sort subjects within a section:
      // 1. Homeroom (First)
      // 2. ARAL (Last)
      // 3. Phil Gov (Mid/Heavy)
      // 4. Heaviest hours first
      sd.sort((a, b) => {
        const aSubj = a[2].toString().toLowerCase();
        const bSubj = b[2].toString().toLowerCase();

        const aHome = aSubj.includes('homeroom') ? 1 : 0;
        const bHome = bSubj.includes('homeroom') ? 1 : 0;
        if (aHome !== bHome) return bHome - aHome;

        const aAral = (aSubj.match(/\baral\b/i) && !aSubj.includes('panlipunan')) ? 1 : 0;
        const bAral = (bSubj.match(/\baral\b/i) && !bSubj.includes('panlipunan')) ? 1 : 0;
        if (aAral !== bAral) return aAral - bAral; // ARAL goes last

        return (parseFloat(b[3]) || 0) - (parseFloat(a[3]) || 0);
      });

      if (!cBooked[section]) cBooked[section] = { 1:[], 2:[], 3:[], 4:[], 5:[] };
      const gradeMatch = section.match(/\b(11|12|7|8|9|10)\b/);
      const grade = gradeMatch ? parseInt(gradeMatch[1], 10) : 7;

      sd.forEach(d => {
        const subject = d[2].toString().trim();
        let hoursLeft = parseFloat(d[3]) || 0;
        const originalHours = hoursLeft;
        
        let candidateSlots = grade >= 11 ? SHS_SLOTS : STANDARD_SLOTS;

        const isPhilGov = subject.toLowerCase().includes('phil gov') || subject.toLowerCase().includes('philippine politics');
        if (isPhilGov) {
          candidateSlots = [
            { in: '7:30 AM',  out: '9:00 AM',  s: 450, e: 540 },
            { in: '9:00 AM',  out: '10:30 AM', s: 540, e: 630 },
            { in: '1:00 PM',  out: '2:30 PM',  s: 780, e: 870 },
            { in: '2:30 PM',  out: '4:00 PM',  s: 870, e: 960 }
          ];
        }

        let prefDays = [1,2,3,4,5];
        const isHomeroom = subject.toLowerCase().includes('homeroom');
        const isAralEnd = subject.match(/\bARAL\b/i) && !subject.toLowerCase().includes('panlipunan');

        if (isAralEnd) {
          candidateSlots = [{ in: '3:00 PM',  out: '4:00 PM',  s: 900, e: 960 }];
        }

        if (isHomeroom) {
          prefDays = [1];
          if (grade >= 11) {
            candidateSlots = [
              { in: '3:30 PM', out: '4:30 PM', s: 930, e: 990 },
              { in: '3:00 PM', out: '4:00 PM', s: 900, e: 960 }
            ];
          } else {
            candidateSlots = [ { in: '7:30 AM', out: '8:30 AM', s: 450, e: 510 } ];
          }
        }

        let slotsAcquired = [];
        let daysUsedForSubject = new Set();

        // Group teachers by specialization for capacity checking
        const specCapacity = {};
        teachers.forEach(t => {
           const spec = t[2].toString().toLowerCase();
           if (spec) {
               if (!specCapacity[spec]) specCapacity[spec] = 0;
               specCapacity[spec]++;
           }
        });

        // Find capacity for this subject
        let maxConcurrent = 1; // Default to at least 1
        let matchedSpec = null;
        Object.keys(specCapacity).forEach(spec => {
           if (subject.toLowerCase().includes(spec) || spec.includes(subject.toLowerCase())) {
               if (!matchedSpec || specCapacity[spec] > maxConcurrent) {
                   maxConcurrent = specCapacity[spec];
                   matchedSpec = spec;
               }
           }
        });

        // Broad search if direct match fails
        if (!matchedSpec) {
            Object.keys(specCapacity).forEach(spec => {
               const words = subject.toLowerCase().split(' ').filter(w => w.length > 3);
               if (words.some(w => spec.includes(w))) {
                   maxConcurrent = Math.max(maxConcurrent, specCapacity[spec]);
               }
            });
        }

        const isFree = (day, slot) => {
          const cConflict = cBooked[section][day].some(b => slot.s < b.e && slot.e > b.s);
          if (cConflict) return false;

          // Check global concurrency against teacher capacity
          let concurrentCount = 0;
          Object.keys(cBooked).forEach(sec => {
             if (sec === section) return;
             // Count how many OTHER sections have a subject matching this capacity booked at this exact time
             // We need to look at outputRows to see what was booked, but cBooked only holds generic times.
             // Wait, cBooked structure is just { s, e }. It doesn't hold the subject name.
             // Let's rely on the outputRows that are accumulating.
          });

          let overlapCount = 0;
          outputRows.forEach(r => {
             // r = [section, subject, m, t, w, th, f, in, out...]
             // check if day matches
             const daysMap = [r[2], r[3], r[4], r[5], r[6]];
             if (daysMap[day - 1]) {
                 const rStart = parseTime(r[7]);
                 const rEnd = parseTime(r[8]);
                 if (slot.s < rEnd && slot.e > rStart) {
                     // It overlaps in time. Does it match our subject type?
                     const rSubj = r[1].toString().toLowerCase();
                     if (matchedSpec && (rSubj.includes(matchedSpec) || matchedSpec.includes(rSubj))) {
                         overlapCount++;
                     } else if (rSubj === subject.toLowerCase()) {
                         overlapCount++;
                     }
                 }
             }
          });

          if (overlapCount >= maxConcurrent) return false;

          return true;
        };

        const bookSlot = (day, slot) => {
          let dur = (slot.e - slot.s) / 60;
          let actualE = slot.e;
          let actualOut = slot.out;
          if (hoursLeft < dur) {
              actualE = slot.s + Math.round(hoursLeft * 60);
              actualOut = formatMinsToTime(actualE);
              dur = hoursLeft;
          }

          cBooked[section][day].push({s: slot.s, e: actualE});
          slotsAcquired.push({ day, in: slot.in, out: actualOut, s: slot.s, e: actualE });
          hoursLeft -= dur;
          daysUsedForSubject.add(day);
        };

        // Horizontal bin-packing to fill gaps
        for (let day of prefDays) {
          if (hoursLeft <= 0) break;
          for (let slot of candidateSlots) {
            if (daysUsedForSubject.has(day) && !isHomeroom) continue;
            if (isFree(day, slot)) { bookSlot(day, slot); break; }
          }
        }

        if (hoursLeft > 0 && !isHomeroom) {
          for (let slot of candidateSlots) { // Fill early slots across all days first
            for (let day of [1,2,3,4,5]) {
              if (hoursLeft <= 0) break;
              if (daysUsedForSubject.has(day)) continue;
              if (isFree(day, slot)) { bookSlot(day, slot); }
            }
            if (hoursLeft <= 0) break;
          }
        }

        if (hoursLeft > 0) {
          unmappedLog.push(`[${term}] ${section} - ${subject}: Could not fit ${hoursLeft}h physically into timetable.`);
        }

        const grouped = {};
        slotsAcquired.forEach(sa => {
          const key = sa.in + '|' + sa.out;
          if (!grouped[key]) grouped[key] = { m:false, t:false, w:false, th:false, f:false, in: sa.in, out: sa.out };
          if (sa.day === 1) grouped[key].m = true;
          if (sa.day === 2) grouped[key].t = true;
          if (sa.day === 3) grouped[key].w = true;
          if (sa.day === 4) grouped[key].th = true;
          if (sa.day === 5) grouped[key].f = true;
        });

        // Autogenerate Suggestions
        let suggestions = [];
        const subjLower = subject.toLowerCase();
        teachers.forEach(t => {
          const spec = t[2].toString().toLowerCase();
          if (spec && subjLower.includes(spec) || spec.includes(subjLower)) {
            suggestions.push(t[1]);
          }
        });

        if (suggestions.length === 0) {
          // Broaden search
          teachers.forEach(t => {
            const spec = t[2].toString().toLowerCase();
            const words = subjLower.split(' ').filter(w => w.length > 3);
            if (words.some(w => spec.includes(w))) suggestions.push(t[1]);
          });
        }

        const suggStr = suggestions.length > 0 ? suggestions.slice(0, 3).join(', ') : 'Any Teacher';

        Object.values(grouped).forEach(g => {
          // A: Section, B: Subject, C-G: MON-FRI, H: Time In, I: Time Out, J: Suggested, K: Assigned Teacher, L: Status, M: Action
          outputRows.push([section, subject, g.m, g.t, g.w, g.th, g.f, g.in, g.out, suggStr, '', '⚪ Pending Assignment', '—']);
        });
      });
    });

    const ts = ss.getSheetByName(term);
    if (ts && outputRows.length > 0) {
      ts.getRange(3, 1, outputRows.length, 13).setValues(outputRows);
      ts.getRange(3, 1, outputRows.length, 13).setVerticalAlignment('middle');
      const actionRule = SpreadsheetApp.newDataValidation().requireValueInList(['—', 'Fix Conflict'], true).build();
      ts.getRange(3, 13, outputRows.length, 1).setDataValidation(actionRule).setHorizontalAlignment('center');

      const cbRule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
      ts.getRange(3, 3, outputRows.length, 5).setDataValidation(cbRule).setHorizontalAlignment('center');
      ts.getRange(3, 8, outputRows.length, 2).setHorizontalAlignment('center');

      ts.getRange(3, 11, outputRows.length, 1).setDataValidation(teacherRule).setBackground(C.tealLight).setFontWeight('bold');
    }
  });

  if (unmappedLog.length > 0) {
    const ui = SpreadsheetApp.getUi();
    ui.alert(
      '⚠️ Timetable Bottleneck',
      'Could not map all hours for sections (timetable full):\n\n' + unmappedLog.join('\n\n'),
      ui.ButtonSet.OK
    );
  } else {
    ss.toast('Student timetables successfully mapped. Ready for Teacher Assignment!', '✅ Ready', 6);
  }
}

// ══════════════════════════════════════════════════════════════
//  SECTION 4: CONFLICT CHECKER ENGINE & AUTO-FIXER
// ══════════════════════════════════════════════════════════════

function runConflictChecker() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rep = ss.getSheetByName(CFG.REPORT);
  if (!rep) return ss.toast('Missing Report Tab. Run Phase 5 Builder first.', '⚠️ Error');

  ss.toast('Scanning all Term schedules...', '🔍 Checking', 3);

  const lastR = rep.getLastRow();
  if (lastR >= 3) {
    rep.getRange(3, 1, lastR - 2, 10).clearDataValidations().clearContent().clearFormat().setBackground(C.white);
  }

  let allConflicts = [];

  CFG.TERMS.forEach(termName => {
    const termSheet = ss.getSheetByName(termName);
    if (!termSheet) return;
    const entries = _buildTermEntries(termSheet, termName);
    const conflicts = _detectConflicts(entries, termName);
    allConflicts = allConflicts.concat(conflicts);
  });

  const ORDER = { ERROR: 0, WARN: 1, INFO: 2 };
  allConflicts.sort((a, b) => (ORDER[a.severity] ?? 2) - (ORDER[b.severity] ?? 2));

  if (!allConflicts.length) {
    rep.getRange('A3:I3').merge().setValue('✅ All clear — no schedule conflicts detected across any term.').setHorizontalAlignment('center').setFontStyle('italic').setFontColor(C.ok).setBackground(C.okBg);
    return;
  }

  const rows = allConflicts.map(c => c.row);
  const payloads = allConflicts.map(c => [c.payload]);

  rep.getRange(3, 1, rows.length, 9).setValues(rows);
  rep.getRange(3, 10, rows.length, 1).setValues(payloads); 

  allConflicts.forEach((c, i) => {
    const sr = 3 + i;
    let bg = C.infoBg, fg = C.info;
    if (c.severity === 'ERROR') { bg = C.errorBg; fg = C.error; }
    if (c.severity === 'WARN')  { bg = C.warnBg; fg = C.warn; }

    rep.getRange(sr, 1, 1, 9).setBackground(bg);
    rep.getRange(sr, 2).setFontWeight('bold').setFontColor(fg); 

    if (c.hasSolution) {
      rep.getRange(sr, 8).setFontWeight('bold').setFontColor(C.ok);
      rep.getRange(sr, 9).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['—', 'Implement'], true).build()).setValue('—').setBackground(C.warnBg).setHorizontalAlignment('center');
    } else {
      rep.getRange(sr, 8).setFontColor(C.muted);
      rep.getRange(sr, 9).setValue('Manual Only').setFontColor(C.muted).setHorizontalAlignment('center');
    }
  });

  rep.getRange(3, 1, rows.length, 9).setVerticalAlignment('middle').setWrap(true);
  rep.getRange(3, 1, rows.length, 5).setHorizontalAlignment('center');
  rep.hideColumns(10); 

  ss.toast(`Scan complete. Found ${allConflicts.length} issues.`, '✅ Done', 5);
}

function applySuggestedFix(e) {
  const sheet = e.range.getSheet();
  const row = e.range.getRow();
  
  const payload = sheet.getRange(row, 10).getValue();
  if (!payload) return;
  
  const [termName, srcRow, newIn, newOut] = payload.split('|');
  const srcSheet = e.source.getSheetByName(termName);
  if (!srcSheet) return;
  
  srcSheet.getRange(parseInt(srcRow, 10), 9).setValue(newIn);
  srcSheet.getRange(parseInt(srcRow, 10), 10).setValue(newOut);
  
  e.range.setValue('✔ Fixed').setBackground(C.okBg).setFontColor(C.ok).clearDataValidations();
  e.source.toast(`Schedule updated in ${termName}. Re-running scan...`, '✅ Applied', 4);
  
  Utilities.sleep(500);
  runConflictChecker(); 
}


// ══════════════════════════════════════════════════════════════
//  SECTION 4.5: REAL-TIME CONFLICT CHECKER
// ══════════════════════════════════════════════════════════════
function checkTeacherConflicts(sheet, teacherName, targetRow) {
  if (!teacherName) {
    sheet.getRange(targetRow, 12).setValue('⚪ Pending Assignment').setBackground(C.white).setFontColor(C.muted);
    return;
  }

  const lastRow = sheet.getLastRow();
  const data = sheet.getRange(3, 1, lastRow - 2, 11).getValues(); // Up to Assigned Teacher

  // Find target row data
  const targetData = data[targetRow - 3];
  const targetDays = [2, 3, 4, 5, 6].filter(d => targetData[d] === true).map(d => ['MON','TUE','WED','THU','FRI'][d-2]);
  const targetStart = parseTime(targetData[7]);
  const targetEnd = parseTime(targetData[8]);
  const targetDuration = calcEffectiveDuration(targetStart, targetEnd);

  let overlaps = [];
  let dailyHours = { 'MON': 0, 'TUE': 0, 'WED': 0, 'THU': 0, 'FRI': 0 };
  let straightHours = { 'MON': { am: 0, pm: 0 }, 'TUE': { am: 0, pm: 0 }, 'WED': { am: 0, pm: 0 }, 'THU': { am: 0, pm: 0 }, 'FRI': { am: 0, pm: 0 } };

  data.forEach((row, idx) => {
    if (row[10] !== teacherName) return; // Not the same teacher

    const start = parseTime(row[7]);
    const end = parseTime(row[8]);
    const dur = calcEffectiveDuration(start, end);
    const activeDays = [2, 3, 4, 5, 6].filter(d => row[d] === true).map(d => ['MON','TUE','WED','THU','FRI'][d-2]);

    activeDays.forEach(day => {
      dailyHours[day] += dur;
      if (start < 720) straightHours[day].am += dur; else straightHours[day].pm += dur;
    });

    // Check overlap if it's not the same exact row
    if (idx !== (targetRow - 3)) {
      const dayOverlap = targetDays.some(d => activeDays.includes(d));
      if (dayOverlap && Math.max(targetStart, start) < Math.min(targetEnd, end)) {
        overlaps.push(`Overlaps with ${row[0]} - ${row[1]}`);
      }
    }
  });

  // Calculate warnings
  let warnings = [];
  if (overlaps.length > 0) warnings.push('🔴 ' + overlaps[0]); // Show first overlap

  targetDays.forEach(day => {
    if (dailyHours[day] / 60 > CFG.DAILY_HARD_HOURS) warnings.push(`🔴 Over 6h limit (${day})`);
    else if (dailyHours[day] / 60 > CFG.DAILY_PREFERRED_HOURS) warnings.push(`🟠 Over 5.5h (${day})`);

    if (straightHours[day].am / 60 > 3) warnings.push(`🟠 >3h straight AM (${day})`);
    if (straightHours[day].pm / 60 > 3) warnings.push(`🟠 >3h straight PM (${day})`);
  });

  const statusCell = sheet.getRange(targetRow, 12);

  if (warnings.length > 0) {
    // Dedup warnings
    warnings = [...new Set(warnings)];
    const hasError = warnings.some(w => w.includes('🔴'));
    statusCell.setValue(warnings.join(', '))
              .setBackground(hasError ? C.errorBg : C.warnBg)
              .setFontColor(hasError ? C.error : C.warn);
  } else {
    statusCell.setValue('✅ Valid Schedule')
              .setBackground(C.okBg)
              .setFontColor(C.ok);
  }
}

function fixTermConflictTab(e, termName) {
  const sheet = e.range.getSheet();
  const row = e.range.getRow();

  const sectionName = sheet.getRange(row, 1).getValue();
  const teacher = sheet.getRange(row, 11).getValue(); // Col K is 11
  const timeIn = sheet.getRange(row, 8).getValue(); // Time In is Col H (8)
  const timeOut = sheet.getRange(row, 9).getValue(); // Time Out is Col I (9)

  if (!teacher) {
    e.range.setValue('—');
    return e.source.toast('Please assign a teacher before attempting to auto-fix conflicts.', '⚠️ Missing Data');
  }

  if (!teacher || !timeIn || !timeOut) {
    e.range.setValue('—');
    return e.source.toast('Missing required schedule data to fix conflict.', '⚠️ Error');
  }

  const start = parseTime(timeIn);
  const end = parseTime(timeOut);
  const durationMins = end - start;

  if (durationMins <= 0) {
    e.range.setValue('—');
    return;
  }

  const entries = _buildTermEntries(sheet, termName);

  // Find which day(s) this row is active
  const days = ['MON','TUE','WED','THU','FRI'];
  const activeDays = days.filter((day, di) => sheet.getRange(row, 3 + di).getValue() === true);

  let fixed = false;

  // Find a single slot that is free across ALL active days
  for (let s = CFG.SCHOOL_START; s + durationMins <= CFG.SCHOOL_END; s += CFG.SLOT_STEP) {
    const end = s + durationMins;
    if (s < CFG.LUNCH_END && end > CFG.LUNCH_START) { s = CFG.LUNCH_END - CFG.SLOT_STEP; continue; }

    let isFreeOnAllDays = true;
    for (let day of activeDays) {
      const byDay = entries.filter(e => e.day === day);
      const hasConflict = byDay.some(entry =>
        s < entry.end && end > entry.start &&
        (entry.teacher === teacher || entry.gradeLevel === sectionName) // using gradeLevel map which is Section in new _buildTermEntries
      );
      if (hasConflict) {
        isFreeOnAllDays = false;
        break;
      }
    }

    if (isFreeOnAllDays) {
      sheet.getRange(row, 8).setValue(formatMinsToTime(s));
      sheet.getRange(row, 9).setValue(formatMinsToTime(end));
      fixed = true;
      e.source.toast(`Fixed overlap for ${teacher}.`, '✅ Fixed', 4);
      checkTeacherConflicts(sheet, teacher, row); // Update live status
      break;
    }
  }

  if (fixed) {
    e.range.setValue('✔ Fixed').setBackground(C.okBg).setFontColor(C.ok).clearDataValidations();
    Utilities.sleep(500);
    // Refresh entries to show it worked
  } else {
    e.range.setValue('—');
    e.source.toast('Could not find a valid free slot to resolve conflict. Manual fix required.', '⚠️ Failed');
  }
}

function _buildTermEntries(src, termName) {
  const last = src.getLastRow();
  if (last < 3) return [];
  const rows = src.getRange(3, 1, last - 2, 11).getValues();
  
  return rows.flatMap((row, idx) => {
    const gradeLevel = row[0], subject = row[1], teacher = row[10];
    const timeIn = row[7], timeOut = row[8];
    if (!teacher || !timeIn || !timeOut) return [];

    const start = parseTime(timeIn);
    const end = parseTime(timeOut);
    if (start >= end) return []; 

    const days = ['MON','TUE','WED','THU','FRI'];
    return days.map((day, di) => ({ day, di }))
      .filter(({ di }) => row[2 + di] === true)
      .map(({ day }) => ({
        term: termName, rowNum: idx + 3,
        teacher, subject, gradeLevel, day, start, end,
        timeStr: `${formatMinsToTime(start)} – ${formatMinsToTime(end)}`,
      }));
  });
}

function _detectConflicts(entries, termName) {
  const conflicts = [];

  entries.forEach(a => {
    if (!(a.start < CFG.LUNCH_END && a.end > CFG.LUNCH_START)) return;
    const slot = _findFreeSlot(entries, a.day, a.teacher, a.gradeLevel, a.end - a.start);
    conflicts.push(_makeConflict(a, null, '🚫 Lunch Block', 'ERROR', slot));
  });

  for (let i = 0; i < entries.length; i++) {
    for (let j = i + 1; j < entries.length; j++) {
      const a = entries[i], b = entries[j];
      if (a.day !== b.day || a.start >= b.end || a.end <= b.start) continue;

      const teacherConflict = a.teacher === b.teacher;
      const cohortConflict  = a.gradeLevel === b.gradeLevel && a.gradeLevel !== '';
      if (!teacherConflict && !cohortConflict) continue;

      const slot = _findFreeSlot(entries, b.day, b.teacher, b.gradeLevel, b.end - b.start);
      const label = teacherConflict ? '⚠️ Teacher Overlap' : '📚 Cohort Overlap';
      conflicts.push(_makeConflict(a, b, label, teacherConflict ? 'ERROR' : 'WARN', slot));
    }
  }

  const map = {};
  entries.forEach(e => {
    if (!map[e.teacher]) map[e.teacher] = { total: 0, daily: { MON:0, TUE:0, WED:0, THU:0, FRI:0 } };
    const durHours = calcEffectiveDuration(e.start, e.end) / 60;
    map[e.teacher].total += durHours;
    map[e.teacher].daily[e.day] += durHours;
  });

  Object.entries(map).forEach(([teacher, stats]) => {
    if (stats.total > CFG.WEEKLY_WARN_HOURS) {
      const isHard = stats.total > CFG.WEEKLY_HARD_HOURS;
      conflicts.push({
        severity: isHard ? 'ERROR' : 'WARN', hasSolution: false, payload: '',
        row: [termName, isHard ? '🔴 Weekly Overload' : '🟠 High Weekly Load', 'ALL', '—', 
              `${stats.total.toFixed(1)} hrs | limit: ${isHard ? CFG.WEEKLY_HARD_HOURS : CFG.WEEKLY_WARN_HOURS}`, 
              teacher, 'All classes', `Reduce by ${(stats.total - (isHard ? CFG.WEEKLY_HARD_HOURS : CFG.WEEKLY_WARN_HOURS)).toFixed(1)}h`, '']
      });
    }
    ['MON','TUE','WED','THU','FRI'].forEach(d => {
      if (stats.daily[d] > CFG.DAILY_WARN_HOURS) {
        const isHard = stats.daily[d] > CFG.DAILY_HARD_HOURS;
        conflicts.push({
          severity: isHard ? 'ERROR' : 'WARN', hasSolution: false, payload: '',
          row: [termName, isHard ? '🔴 Daily Overload' : '🟠 High Daily Load', d, '—', 
                `${stats.daily[d].toFixed(1)} hrs | limit: ${isHard ? CFG.DAILY_HARD_HOURS : CFG.DAILY_WARN_HOURS}`, 
                teacher, 'Daily load', `Reduce by ${(stats.daily[d] - (isHard ? CFG.DAILY_HARD_HOURS : CFG.DAILY_WARN_HOURS)).toFixed(1)}h`, '']
        });
      }
    });
  });

  return conflicts;
}

function _findFreeSlot(allEntries, day, teacher, gradeLevel, durationMins) {
  const byDay = allEntries.filter(e => e.day === day);
  for (let s = CFG.SCHOOL_START; s + durationMins <= CFG.SCHOOL_END; s += CFG.SLOT_STEP) {
    const end = s + durationMins;
    if (s < CFG.LUNCH_END && end > CFG.LUNCH_START) { s = CFG.LUNCH_END - CFG.SLOT_STEP; continue; }
    
    const hasConflict = byDay.some(e => s < e.end && end > e.start && 
      (e.teacher === teacher || (e.gradeLevel === gradeLevel && gradeLevel !== '')));
    
    if (!hasConflict) return { start: s, end };
  }
  return null;
}

function _makeConflict(a, b, type, severity, slot) {
  const sugg = slot ? `Move to ${formatMinsToTime(slot.start)} – ${formatMinsToTime(slot.end)}` : 'Manual Fix Req.';
  const payload = slot ? `${a.term}|${b ? b.rowNum : a.rowNum}|${formatMinsToTime(slot.start)}|${formatMinsToTime(slot.end)}` : '';
  
  if (!b) { 
    return { severity, hasSolution: !!slot, payload, row: [
      a.term, type, a.day, a.gradeLevel, `${a.timeStr} ⚡ LUNCH`, 
      `${a.teacher}\n(Row ${a.rowNum})`, a.subject, sugg, ''
    ]};
  }
  
  const grade = a.gradeLevel === b.gradeLevel ? a.gradeLevel : `${a.gradeLevel} / ${b.gradeLevel}`;
  return { severity, hasSolution: !!slot, payload, row: [
    a.term, type, a.day, grade, `${a.timeStr} ⚡ ${b.timeStr}`, 
    `${a.teacher}\n${a.subject} (Row ${a.rowNum})`, `${b.teacher}\n${b.subject} (Row ${b.rowNum})`, sugg, ''
  ]};
}

// ══════════════════════════════════════════════════════════════
//  SECTION 5: DASHBOARD ENGINES (Teacher & Section)
// ══════════════════════════════════════════════════════════════

function updateDashboardUI(dash, ss, pane = 1) {
  const colOffset = pane === 1 ? 0 : 10; // Pane 1 is col 1, Pane 2 is col 11 (K)
  const prefix = pane === 1 ? '' : '1';

  const teacher = dash.getRange(2, 2 + colOffset).getValue();
  const day     = dash.getRange(3, 2 + colOffset).getValue();
  const term    = dash.getRange(4, 2 + colOffset).getValue();

  const maxR = dash.getMaxRows();
  if (maxR > 5) dash.getRange(6, 1 + colOffset, maxR - 5, 9).clearContent().clearDataValidations().clearFormat();
  
  dash.getRange(4, 3 + colOffset).setValue('');
  dash.getRange(4, 5 + colOffset).setValue('Select options...').setBackground(C.tealLight).setFontColor(C.teal);

  if (!teacher || !term) return;

  const src = ss.getSheetByName(term);
  if (!src) return;

  const dayMap = { MON: 2, TUE: 3, WED: 4, THU: 5, FRI: 6 };
  const lastRow = src.getLastRow();
  if (lastRow < 3) return;
  const rows = src.getRange(3, 1, lastRow - 2, 11).getValues();
  
  let output = [];
  let totalMins = 0;

  rows.forEach(row => {
    if (row[10] !== teacher) return; // Column K
    
    const start = parseTime(row[7]);
    const end   = parseTime(row[8]);
    const dur   = calcEffectiveDuration(start, end); 

    if (day === 'ALL WEEK') {
      let daysOn = 0;
      for (let d = 2; d <= 6; d++) { if (row[d] === true) daysOn++; }
      totalMins += dur * daysOn;
      output.push([row[1], row[0], row[7], row[8], row[2], row[3], row[4], row[5], row[6]]);
    } else {
      if (row[dayMap[day]] !== true) return;
      totalMins += dur;
      output.push([row[1], row[0], row[7], row[8], '', '', '', '', '']);
    }
  });

  // Sort Chronologically
  output.sort((a, b) => parseTime(a[2]) - parseTime(b[2]));

  const totalHours = totalMins / 60;
  dash.getRange(4, 3 + colOffset).setValue(totalHours);

  let statusText = '✅ Normal Load';
  let statusBg = C.okBg; let statusFg = C.ok;

  const isWeekly = day === 'ALL WEEK';
  const hardLimit = isWeekly ? CFG.WEEKLY_HARD_HOURS : CFG.DAILY_HARD_HOURS;
  const warnLimit = isWeekly ? CFG.WEEKLY_WARN_HOURS : CFG.DAILY_WARN_HOURS;

  if (totalHours > hardLimit) {
    statusText = '🔴 Overloaded!'; statusBg = C.errorBg; statusFg = C.error;
  } else if (totalHours > warnLimit) {
    statusText = '🟠 High Load'; statusBg = C.warnBg; statusFg = C.warn;
  } else if (totalHours === 0) {
    statusText = '⚪ No Load'; statusBg = C.navyLight; statusFg = C.muted;
  }

  dash.getRange(4, 3 + colOffset).setBackground(statusBg).setFontColor(statusFg);
  dash.getRange(4, 5 + colOffset).setValue(statusText).setBackground(statusBg).setFontColor(statusFg);

  if (!output.length) {
    dash.getRange(6, 1 + colOffset, 1, 9).merge().setValue('No schedule found for these criteria.').setFontStyle('italic').setFontColor(C.muted).setHorizontalAlignment('center');
    return;
  }

  dash.getRange(6, 1 + colOffset, output.length, 9).setValues(output);
  for (let r = 0; r < output.length; r++) {
    dash.getRange(6 + r, 1 + colOffset, 1, 9).setBackground(r % 2 === 0 ? C.white : C.rowAlt).setFontColor(C.body);
  }
  dash.getRange(6, 2 + colOffset, output.length, 1).setHorizontalAlignment('center');
  dash.getRange(6, 3 + colOffset, output.length, 2).setNumberFormat('h:mm AM/PM').setHorizontalAlignment('center');
  if (isWeekly) {
    const rule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
    dash.getRange(6, 5 + colOffset, output.length, 5).setDataValidation(rule).setHorizontalAlignment('center');
  }
}

function updateSectionDashboardUI(dash, ss) {
  const section = dash.getRange('B2').getValue();
  const day     = dash.getRange('B3').getValue();
  const term    = dash.getRange('B4').getValue();

  const maxR = dash.getMaxRows();
  if (maxR > 5) dash.getRange(6, 1, maxR - 5, 9).clearContent().clearDataValidations().clearFormat();
  
  dash.getRange('C4').setValue('');
  dash.getRange('E4').setValue('Select options...').setBackground(C.tealLight).setFontColor(C.teal);

  if (!section || !term) return;

  const src = ss.getSheetByName(term);
  if (!src) return;

  const dayMap = { MON: 2, TUE: 3, WED: 4, THU: 5, FRI: 6 };
  const lastRow = src.getLastRow();
  if (lastRow < 3) return;
  const rows = src.getRange(3, 1, lastRow - 2, 11).getValues();
  
  let output = [];
  let totalMins = 0;

  rows.forEach(row => {
    if (row[0] !== section) return; // Match Section
    
    const start = parseTime(row[7]);
    const end   = parseTime(row[8]);
    const dur   = calcEffectiveDuration(start, end); 

    if (day === 'ALL WEEK') {
      let daysOn = 0;
      for (let d = 2; d <= 6; d++) { if (row[d] === true) daysOn++; }
      totalMins += dur * daysOn;
      output.push([row[1], row[10] || 'Pending', row[7], row[8], row[2], row[3], row[4], row[5], row[6]]);
    } else {
      if (row[dayMap[day]] !== true) return;
      totalMins += dur;
      output.push([row[1], row[10] || 'Pending', row[7], row[8], '', '', '', '', '']);
    }
  });

  // Sort Chronologically
  output.sort((a, b) => parseTime(a[2]) - parseTime(b[2]));

  const totalHours = totalMins / 60;
  dash.getRange('C4').setValue(totalHours);

  let statusText = '✅ Balanced Schedule';
  let statusBg = C.okBg; let statusFg = C.ok;

  const isWeekly = day === 'ALL WEEK';
  // Student limits tend to be slightly higher than teacher limits before burnout
  const hardLimit = isWeekly ? 35 : 7.5; 
  const warnLimit = isWeekly ? 32 : 6.5;

  if (totalHours > hardLimit) {
    statusText = '🔴 Heavy Academic Load'; statusBg = C.errorBg; statusFg = C.error;
  } else if (totalHours > warnLimit) {
    statusText = '🟠 High Contact Hours'; statusBg = C.warnBg; statusFg = C.warn;
  } else if (totalHours === 0) {
    statusText = '⚪ No Classes'; statusBg = C.navyLight; statusFg = C.muted;
  }

  dash.getRange('C4').setBackground(statusBg).setFontColor(statusFg);
  dash.getRange('E4').setValue(statusText).setBackground(statusBg).setFontColor(statusFg);

  if (!output.length) {
    dash.getRange('A6:I6').merge().setValue('No schedule found for this cohort.').setFontStyle('italic').setFontColor(C.muted).setHorizontalAlignment('center');
    return;
  }

  dash.getRange(6, 1, output.length, 9).setValues(output);
  for (let r = 0; r < output.length; r++) {
    dash.getRange(6 + r, 1, 1, 9).setBackground(r % 2 === 0 ? C.white : C.rowAlt).setFontColor(C.body);
  }
  dash.getRange(6, 2, output.length, 1).setHorizontalAlignment('center');
  dash.getRange(6, 3, output.length, 2).setNumberFormat('h:mm AM/PM').setHorizontalAlignment('center');
  if (isWeekly) {
    const rule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
    dash.getRange(6, 5, output.length, 5).setDataValidation(rule).setHorizontalAlignment('center');
  }
}

// ══════════════════════════════════════════════════════════════
//  SECTION 6: DATA SYNC (Cascading Deletes & Dropdowns)
// ══════════════════════════════════════════════════════════════

function cleanOrphanedData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const secSheet = ss.getSheetByName(CFG.SECTION_ENROLL);
  const subSheet = ss.getSheetByName(CFG.SUBJECT_LOAD);
  const teachAssignSheet = ss.getSheetByName(CFG.TEACHER_ASSIGN);
  const teachEnrollSheet = ss.getSheetByName(CFG.TEACHER_ENROLL);

  if (!secSheet || !subSheet || !teachAssignSheet || !teachEnrollSheet) return;

  const validSections = new Set(secSheet.getRange('B3:B').getValues().flat().filter(String));
  const validSubjects = new Set(subSheet.getRange('C3:C').getValues().flat().filter(String));
  const validTeachers = new Set(teachEnrollSheet.getRange('B3:B').getValues().flat().filter(String));

  const subLastRow = Math.max(3, subSheet.getLastRow());
  const subRange = subSheet.getRange('B3:B' + subLastRow);
  const subValues = subRange.getValues();
  let subChanged = false;
  
  subValues.forEach(row => {
    if (row[0] && !validSections.has(row[0])) { row[0] = ''; subChanged = true; }
  });
  if (subChanged) subRange.setValues(subValues);

  const teachLastRow = Math.max(3, teachAssignSheet.getLastRow());
  const teachRange = teachAssignSheet.getRange('A3:C' + teachLastRow); 
  const teachValues = teachRange.getValues();
  let teachChanged = false;

  teachValues.forEach(row => {
    if (row[0] && !validTeachers.has(row[0])) { row[0] = ''; teachChanged = true; }
    if (row[1] && !validSubjects.has(row[1])) { row[1] = ''; teachChanged = true; }
    if (row[2] && !validSections.has(row[2])) { row[2] = ''; teachChanged = true; }
  });
  
  if (teachChanged) teachRange.setValues(teachValues);
}

function updateSectionDropdowns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const secSheet = ss.getSheetByName(CFG.SECTION_ENROLL);
  const subSheet = ss.getSheetByName(CFG.SUBJECT_LOAD);
  const secDashSheet = ss.getSheetByName(CFG.SECTION_DASH); // NEW
  
  if (!secSheet) return;
  const rule = SpreadsheetApp.newDataValidation().requireValueInRange(secSheet.getRange('B3:B1000'), true).build();
  
  if (subSheet) subSheet.getRange('B3:B1000').setDataValidation(rule);
  if (secDashSheet) secDashSheet.getRange('B2').setDataValidation(rule); // Sync to Dashboard
}

function updateTeacherNameDropdowns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const enrollSheet = ss.getSheetByName(CFG.TEACHER_ENROLL);
  const assignSheet = ss.getSheetByName(CFG.TEACHER_ASSIGN);
  const dashSheet   = ss.getSheetByName(CFG.DASHBOARD); 

  if (!enrollSheet) return;
  const rule = SpreadsheetApp.newDataValidation().requireValueInRange(enrollSheet.getRange('B3:B1000'), true).build();
  
  if (assignSheet) assignSheet.getRange('A3:A1000').setDataValidation(rule);
  if (dashSheet) {
    dashSheet.getRange('B2').setDataValidation(rule);
    dashSheet.getRange('L2').setDataValidation(rule);
  }
}

function updateTeacherDropdownOptions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const subSheet = ss.getSheetByName(CFG.SUBJECT_LOAD);
  const teachSheet = ss.getSheetByName(CFG.TEACHER_ASSIGN);
  if (!subSheet || !teachSheet) return;
  
  const subjectRule = SpreadsheetApp.newDataValidation().requireValueInRange(subSheet.getRange('C3:C1000'), true).build();
  teachSheet.getRange('B3:B1000').setDataValidation(subjectRule);
  
  const sectionRule = SpreadsheetApp.newDataValidation().requireValueInRange(subSheet.getRange('B3:B1000'), true).build();
  teachSheet.getRange('C3:C1000').setDataValidation(sectionRule);
}

// ══════════════════════════════════════════════════════════════
//  SECTION 7: TAB BUILDERS (Phases 1-6)
// ══════════════════════════════════════════════════════════════

function buildPhase1Tabs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  let secSheet = ss.getSheetByName(CFG.SECTION_ENROLL) || ss.insertSheet(CFG.SECTION_ENROLL, 0);
  secSheet.clear();
  secSheet.getRange('A1:B1').merge().setValue('🏫 SECTION ENROLLMENT (COHORTS)').setFontWeight('bold').setFontSize(12).setBackground(C.navyDark).setFontColor(C.white).setHorizontalAlignment('center').setVerticalAlignment('middle');
  secSheet.getRange('A2:B2').setValues([['Grade Level', 'Section Name']]).setFontWeight('bold').setFontSize(10).setBackground(C.navy).setFontColor(C.white).setHorizontalAlignment('center').setVerticalAlignment('middle');
  secSheet.setColumnWidth(1, 150); secSheet.setColumnWidth(2, 250);
  secSheet.getRange('A3:B1000').setFontColor(C.body).setVerticalAlignment('middle');
  secSheet.getRange('A3:A1000').setHorizontalAlignment('center');
  secSheet.setFrozenRows(2);

  let subSheet = ss.getSheetByName(CFG.SUBJECT_LOAD) || ss.insertSheet(CFG.SUBJECT_LOAD, 1);
  subSheet.clear();
  subSheet.getRange('A1:D1').merge().setValue('📚 SUBJECT LOADING (CURRICULUM DEMAND)').setFontWeight('bold').setFontSize(12).setBackground(C.navyDark).setFontColor(C.white).setHorizontalAlignment('center').setVerticalAlignment('middle');
  subSheet.getRange('A2:D2').setValues([['Term', 'Section', 'Subject', 'Weekly Class Hours']]).setFontWeight('bold').setFontSize(10).setBackground(C.teal).setFontColor(C.white).setHorizontalAlignment('center').setVerticalAlignment('middle');
  subSheet.setColumnWidth(1, 120); subSheet.setColumnWidth(2, 200); subSheet.setColumnWidth(3, 250); subSheet.setColumnWidth(4, 150);
  subSheet.getRange('A3:D1000').setFontColor(C.body).setVerticalAlignment('middle');
  subSheet.getRange('D3:D1000').setNumberFormat('0.0').setHorizontalAlignment('center');
  subSheet.getRange('A3:A1000').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(CFG.TERMS, true).build()).setHorizontalAlignment('center');
  subSheet.setFrozenRows(2);

  updateSectionDropdowns();
  ss.toast('Phase 1 tabs ready.', '✅ Complete', 4);
}

function buildPhase2TeacherTabs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  let enrollSheet = ss.getSheetByName(CFG.TEACHER_ENROLL) || ss.insertSheet(CFG.TEACHER_ENROLL, 2);
  enrollSheet.clear();
  enrollSheet.getRange('A1:C1').merge().setValue('👨‍🏫 TEACHER ENROLLMENT (FACULTY ROSTER)').setFontWeight('bold').setFontSize(12).setBackground(C.navyDark).setFontColor(C.white).setHorizontalAlignment('center').setVerticalAlignment('middle');
  enrollSheet.getRange('A2:C2').setValues([['Department / Rank', 'Teacher Name', 'Major of Specialization']]).setFontWeight('bold').setFontSize(10).setBackground(C.navy).setFontColor(C.white).setHorizontalAlignment('center').setVerticalAlignment('middle');
  enrollSheet.setColumnWidth(1, 180); enrollSheet.setColumnWidth(2, 250); enrollSheet.setColumnWidth(3, 220); 
  enrollSheet.getRange('A3:C1000').setFontColor(C.body).setVerticalAlignment('middle');
  enrollSheet.getRange('A3:A1000').setHorizontalAlignment('center');
  enrollSheet.getRange('C3:C1000').setHorizontalAlignment('center');
  enrollSheet.setFrozenRows(2);

  let assignSheet = ss.getSheetByName(CFG.TEACHER_ASSIGN) || ss.insertSheet(CFG.TEACHER_ASSIGN, 3);
  assignSheet.clear();
  assignSheet.getRange('A1:C1').merge().setValue('👤 FACULTY ASSIGNMENTS & PREFERENCES').setFontWeight('bold').setFontSize(12).setBackground(C.navyDark).setFontColor(C.white).setHorizontalAlignment('center').setVerticalAlignment('middle');
  assignSheet.getRange('A2:C2').setValues([['Teacher Name', 'Assigned Subject', 'Assigned Section']]).setFontWeight('bold').setFontSize(10).setBackground(C.navy).setFontColor(C.white).setHorizontalAlignment('center').setVerticalAlignment('middle');
  assignSheet.setColumnWidth(1, 250); assignSheet.setColumnWidth(2, 250); assignSheet.setColumnWidth(3, 200);
  assignSheet.getRange('A3:C1000').setFontColor(C.body).setVerticalAlignment('middle');
  assignSheet.setFrozenRows(2);

  updateTeacherNameDropdowns();
  updateTeacherDropdownOptions();
  ss.toast('Teacher Enrollment and Assignment tabs initialized.', '✅ Phase 2 Complete', 5);
}

function buildPhase3TermTabs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  CFG.TERMS.forEach((term, index) => {
    let sheet = ss.getSheetByName(term) || ss.insertSheet(term, 4 + index);
    sheet.clear();
    sheet.getRange('A1:L1').merge().setValue(`📅 ${term.toUpperCase()} MASTER SCHEDULE`).setFontWeight('bold').setFontSize(12).setBackground(C.navyDark).setFontColor(C.white).setHorizontalAlignment('center').setVerticalAlignment('middle');
    sheet.setRowHeight(1, 40);

    const headers = ['Section', 'Subject', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Time In', 'Time Out', 'Suggested Teachers', 'Assigned Teacher', 'Status / Conflicts', 'Action'];
    sheet.getRange('A2:M2').setValues([headers]).setFontWeight('bold').setFontSize(10).setBackground(C.teal).setFontColor(C.white).setHorizontalAlignment('center').setVerticalAlignment('middle');
    sheet.setRowHeight(2, 30);

    sheet.setColumnWidth(1, 150); sheet.setColumnWidth(2, 200);
    sheet.setColumnWidths(3, 5, 60); sheet.setColumnWidths(8, 2, 85);
    sheet.setColumnWidth(10, 180); sheet.setColumnWidth(11, 200); sheet.setColumnWidth(12, 280); sheet.setColumnWidth(13, 120);

    sheet.getRange('A3:M1000').setFontColor(C.body).setVerticalAlignment('middle');

    const actionRule = SpreadsheetApp.newDataValidation().requireValueInList(['—', 'Fix Conflict'], true).build();
    sheet.getRange('M3:M1000').setDataValidation(actionRule).setValue('—').setHorizontalAlignment('center');

    const cbRule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
    sheet.getRange('C3:G1000').setDataValidation(cbRule).setHorizontalAlignment('center');

    sheet.getRange('H3:I1000').setHorizontalAlignment('center').setNumberFormat('h:mm AM/PM');

    sheet.getRange('A3:A1000').setHorizontalAlignment('center'); 
    sheet.getRange('B3:B1000').setHorizontalAlignment('left');
    sheet.getRange('J3:J1000').setFontColor(C.muted).setFontStyle('italic').setWrap(true);
    sheet.setFrozenRows(2);
  });

  ss.toast('Term schedules successfully built.', '✅ Phase 3 Complete', 5);
}

function buildPhase4Dashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let dash = ss.getSheetByName(CFG.DASHBOARD) || ss.insertSheet(CFG.DASHBOARD, 1); 
  dash.clear();

  dash.getRange('A1:S1').merge().setValue('🏫 FACULTY LOADING — DUAL TEACHER COMPARISON VIEWER').setFontWeight('bold').setFontSize(14).setBackground(C.navyDark).setFontColor(C.white).setHorizontalAlignment('center').setVerticalAlignment('middle');
  dash.setRowHeight(1, 46);

  const lbl = (cell, text) => dash.getRange(cell).setValue(text).setFontWeight('bold').setHorizontalAlignment('right').setFontColor(C.navy);
  const inputBox = (a1, bg) => dash.getRange(a1).setBackground(bg).setBorder(true,true,true,true,null,null,C.border,SpreadsheetApp.BorderStyle.SOLID).setFontColor(C.navy);

  // Pane 1 (Left)
  lbl('A2', '👤 Teacher 1'); lbl('A3', '📅 Day Filter'); lbl('A4', '📑 Term');
  dash.getRange('D4').setValue('Status').setFontColor(C.muted).setFontSize(10).setHorizontalAlignment('right');
  inputBox('B2', C.navyLight); inputBox('B3', C.navyLight); inputBox('B4', C.navyLight);
  inputBox('C4', C.tealLight).setNumberFormat('0.0 "hrs"').setFontWeight('bold'); 
  inputBox('E4', C.tealLight).setFontWeight('bold').setFontSize(10).setHorizontalAlignment('center'); 

  dash.getRange('B3').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['ALL WEEK','MON','TUE','WED','THU','FRI'], true).build()).setValue('ALL WEEK');
  dash.getRange('B4').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(CFG.TERMS, true).build()).setValue(CFG.TERMS[0]);

  dash.getRange('A5:I5').setValues([['Subject','Grade Level','Time In','Time Out','MON','TUE','WED','THU','FRI']]).setFontWeight('bold').setFontSize(10).setBackground(C.navy).setFontColor(C.white).setHorizontalAlignment('center').setVerticalAlignment('middle');

  // Spacer
  dash.setColumnWidth(10, 20);
  dash.getRange('J2:J1000').setBackground(C.navyLight);

  // Pane 2 (Right)
  lbl('K2', '👤 Teacher 2'); lbl('K3', '📅 Day Filter'); lbl('K4', '📑 Term');
  dash.getRange('N4').setValue('Status').setFontColor(C.muted).setFontSize(10).setHorizontalAlignment('right');
  inputBox('L2', C.navyLight); inputBox('L3', C.navyLight); inputBox('L4', C.navyLight);
  inputBox('M4', C.tealLight).setNumberFormat('0.0 "hrs"').setFontWeight('bold');
  inputBox('O4', C.tealLight).setFontWeight('bold').setFontSize(10).setHorizontalAlignment('center');

  dash.getRange('L3').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['ALL WEEK','MON','TUE','WED','THU','FRI'], true).build()).setValue('ALL WEEK');
  dash.getRange('L4').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(CFG.TERMS, true).build()).setValue(CFG.TERMS[0]);

  dash.getRange('K5:S5').setValues([['Subject','Grade Level','Time In','Time Out','MON','TUE','WED','THU','FRI']]).setFontWeight('bold').setFontSize(10).setBackground(C.teal).setFontColor(C.white).setHorizontalAlignment('center').setVerticalAlignment('middle');

  dash.setRowHeights(2, 3, 28);
  dash.setRowHeight(5, 30);
  
  dash.setColumnWidth(1, 260); dash.setColumnWidth(2, 130); dash.setColumnWidths(3, 2, 110); dash.setColumnWidths(5, 5, 52);
  dash.setColumnWidth(11, 260); dash.setColumnWidth(12, 130); dash.setColumnWidths(13, 2, 110); dash.setColumnWidths(15, 5, 52);

  dash.getRange('C6:D').setNumberFormat('h:mm AM/PM');
  dash.getRange('M6:N').setNumberFormat('h:mm AM/PM');

  dash.setFrozenRows(5); dash.showColumns(5, 5); dash.showColumns(15, 5);
  updateTeacherNameDropdowns();

  ss.toast('Dual Teacher Dashboard fully constructed.', '✅ Phase 4 Complete', 5);
}

function buildPhase5ConflictReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let rep = ss.getSheetByName(CFG.REPORT) || ss.insertSheet(CFG.REPORT, 2);
  rep.clear();

  rep.getRange('A1:I1').merge().setValue('🔍 INTELLIGENT CONFLICT & ALIGNMENT CHECKER').setFontWeight('bold').setFontSize(14).setBackground(C.navyDark).setFontColor(C.white).setHorizontalAlignment('center').setVerticalAlignment('middle');
  rep.setRowHeight(1, 46);

  const headers = ['Term', 'Conflict Type', 'Day', 'Grade Level / Cohort', 'Time Overlap', 'Teacher 1 Details', 'Teacher 2 Details', 'Suggested Solution', 'Action'];
  rep.getRange('A2:I2').setValues([headers]).setFontWeight('bold').setFontSize(10).setBackground(C.navy).setFontColor(C.white).setHorizontalAlignment('center').setVerticalAlignment('middle');
  rep.setRowHeight(2, 30);

  rep.setColumnWidth(1, 80); rep.setColumnWidth(2, 140); rep.setColumnWidth(3, 60); rep.setColumnWidth(4, 150); rep.setColumnWidth(5, 160); rep.setColumnWidths(6, 2, 180); rep.setColumnWidth(8, 240); rep.setColumnWidth(9, 110); 
  rep.getRange('A3:I10').merge().setValue('Run the Conflict Checker from the "Faculty Tools" menu to scan the Terms.').setFontStyle('italic').setFontColor(C.muted).setHorizontalAlignment('center').setVerticalAlignment('middle');
  
  rep.setFrozenRows(2);
  ss.toast('Conflict Report tab constructed.', '✅ Phase 5 Complete', 5);
}

function buildPhase6SectionDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let dash = ss.getSheetByName(CFG.SECTION_DASH) || ss.insertSheet(CFG.SECTION_DASH, 2); 
  dash.clear();

  dash.getRange('A1:I1').merge().setValue('🎓 COHORT VIEWER — SECTION SCHEDULE')
    .setFontWeight('bold').setFontSize(14).setBackground(C.navyDark).setFontColor(C.white)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  dash.setRowHeight(1, 46);

  const lbl = (cell, text) => dash.getRange(cell).setValue(text).setFontWeight('bold').setHorizontalAlignment('right').setFontColor(C.navy);
  lbl('A2', '👥 Section'); lbl('A3', '📅 Day Filter'); lbl('A4', '📑 Term');
  dash.getRange('D4').setValue('Status').setFontColor(C.muted).setFontSize(10).setHorizontalAlignment('right');

  const inputBox = (a1, bg) => dash.getRange(a1).setBackground(bg).setBorder(true,true,true,true,null,null,C.border,SpreadsheetApp.BorderStyle.SOLID).setFontColor(C.navy);
  inputBox('B2', C.navyLight); inputBox('B3', C.navyLight); inputBox('B4', C.navyLight);
  inputBox('C4', C.tealLight).setNumberFormat('0.0 "hrs"').setFontWeight('bold'); 
  inputBox('E4', C.tealLight).setFontWeight('bold').setFontSize(10).setHorizontalAlignment('center'); 
  dash.setRowHeights(2, 3, 28);

  dash.getRange('B3').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['ALL WEEK','MON','TUE','WED','THU','FRI'], true).build()).setValue('ALL WEEK');
  dash.getRange('B4').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(CFG.TERMS, true).build()).setValue(CFG.TERMS[0]);

  // Note the layout change: Subject and Teacher
  dash.getRange('A5:I5').setValues([['Subject','Teacher','Time In','Time Out','MON','TUE','WED','THU','FRI']]).setFontWeight('bold').setFontSize(10).setBackground(C.teal).setFontColor(C.white).setHorizontalAlignment('center').setVerticalAlignment('middle');
  dash.setRowHeight(5, 30);
  
  dash.setColumnWidth(1, 260); dash.setColumnWidth(2, 130); dash.setColumnWidths(3, 2, 110); dash.setColumnWidths(5, 5, 52);
  dash.getRange('C6:D').setNumberFormat('h:mm AM/PM');
  dash.setFrozenRows(5); dash.showColumns(5, 5);
  
  updateSectionDropdowns();

  ss.toast('Section Dashboard fully constructed.', '✅ Phase 6 Complete', 5);
}

// ══════════════════════════════════════════════════════════════
//  SECTION 8: TIME & MATH UTILITIES
// ══════════════════════════════════════════════════════════════

function calcEffectiveDuration(startMins, endMins) {
  if (endMins <= startMins) return 0;
  let mins = endMins - startMins;
  const os = Math.max(startMins, CFG.LUNCH_START);
  const oe = Math.min(endMins,   CFG.LUNCH_END);
  if (oe > os) mins -= (oe - os);
  return Math.max(0, mins);
}

function parseTime(v) {
  if (!v && v !== 0) return 0;
  if (v instanceof Date) return v.getHours() * 60 + v.getMinutes();
  if (typeof v === 'number') return Math.round(v * 1440) % 1440;
  const s = v.toString().trim().toUpperCase();
  const ap = s.match(/^(\d{1,2}):(\d{2})\s*(AM|PM)$/);
  if (ap) {
    let h = parseInt(ap[1], 10), m = parseInt(ap[2], 10);
    if (ap[3] === 'PM' && h !== 12) h += 12;
    if (ap[3] === 'AM' && h === 12) h  =  0;
    return h * 60 + m;
  }
  const h24 = s.match(/^(\d{1,2}):(\d{2})$/);
  if (h24) return parseInt(h24[1], 10) * 60 + parseInt(h24[2], 10);
  return 0;
}

function formatMinsToTime(mins) {
  const h = Math.floor(mins / 60);
  const m = mins % 60;
  return `${h % 12 || 12}:${String(m).padStart(2, '0')} ${h >= 12 ? 'PM' : 'AM'}`;
}

// ══════════════════════════════════════════════════════════════
//  SECTION 9: PDF GENERATION (Class Programs)
// ══════════════════════════════════════════════════════════════

function generateScheduleUI() {
  const htmlOutput = HtmlService.createHtmlOutput(`
    <div style="font-family: Arial, sans-serif; padding: 10px;">
      <h4>Generate Schedule PDF</h4>
      <label><b>Mode:</b></label><br>
      <select id="mode" style="width: 100%; padding: 5px; margin-bottom: 10px;">
        <option value="Teacher">Per Teacher</option>
        <option value="Section">Per Section</option>
      </select>

      <label><b>Term:</b></label><br>
      <select id="term" style="width: 100%; padding: 5px; margin-bottom: 10px;">
        ${CFG.TERMS.map(t => `<option value="${t}">${t}</option>`).join('')}
      </select>

      <label><b>Name (Exact text):</b></label><br>
      <input type="text" id="targetName" placeholder="e.g. MICHELLE JANE GACUSAN or 11-Lone" style="width: 100%; padding: 5px; margin-bottom: 15px;">

      <button onclick="generate()" style="background-color: #1a3a6e; color: white; border: none; padding: 8px 16px; cursor: pointer; width: 100%;">Generate PDF</button>

      <p id="status" style="margin-top: 15px; font-size: 12px; color: #b71c1c;"></p>

      <script>
        function generate() {
          const mode = document.getElementById('mode').value;
          const term = document.getElementById('term').value;
          const targetName = document.getElementById('targetName').value;

          if (!targetName) {
            document.getElementById('status').innerText = 'Please enter a name or section.';
            return;
          }

          document.getElementById('status').innerText = 'Generating PDF... please wait. This may take up to 30 seconds.';
          document.getElementById('status').style.color = '#00796b';

          google.script.run.withSuccessHandler(url => {
            if (url.startsWith('Error:')) {
               document.getElementById('status').innerText = url;
               document.getElementById('status').style.color = '#b71c1c';
            } else {
               document.getElementById('status').innerHTML = '<a href="' + url + '" target="_blank" style="color: #1a3a6e; font-weight: bold;">Click here to download/view your PDF</a>';
            }
          }).processPDFGeneration(mode, term, targetName);
        }
      </script>
    </div>
  `).setWidth(300).setHeight(320);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, '🖨️ PDF Exporter');
}

function processPDFGeneration(mode, termName, targetName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const termSheet = ss.getSheetByName(termName);
    if (!termSheet) return 'Error: Term sheet not found.';

    const data = termSheet.getRange(3, 1, Math.max(1, termSheet.getLastRow() - 2), 12).getValues();

    // Filter matching rows
    let rows = [];
    if (mode === 'Teacher') {
      rows = data.filter(r => r[10] && r[10].toString().toLowerCase() === targetName.toLowerCase()); // Assigned Teacher is col 10
    } else {
      rows = data.filter(r => r[0] && r[0].toString().toLowerCase() === targetName.toLowerCase()); // Section is col 0
    }

    if (!rows.length) return 'Error: No schedule found for ' + targetName;

    return buildPDFTemplateSheetAndExport(mode, targetName, rows);
  } catch (err) {
    return 'Error: ' + err.toString();
  }
}

function buildPDFTemplateSheetAndExport(mode, targetName, rows) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Create a temporary sheet
  const tempName = 'TEMP_PDF_' + new Date().getTime();
  const ts = ss.insertSheet(tempName);

  try {
    // 1. Setup Headers
    if (mode === 'Teacher') {
      ts.getRange('A1:G1').merge().setValue('Name of Teacher: ' + targetName.toUpperCase()).setFontWeight('bold').setFontSize(14);
      // Try to find specialization from Phase 2
      const teachList = ss.getSheetByName(CFG.TEACHER_ENROLL);
      let spec = 'UNKNOWN';
      if (teachList) {
         const tData = teachList.getRange(3,2,teachList.getLastRow(), 2).getValues();
         const match = tData.find(r => r[0].toString().toLowerCase() === targetName.toLowerCase());
         if (match) spec = match[1];
      }
      ts.getRange('A2:G2').merge().setValue('Specialization: ' + spec).setFontWeight('bold').setFontSize(14);
    } else {
      ts.getRange('A1:G1').merge().setValue('Grade Level: ' + targetName).setFontWeight('bold').setFontSize(14);
      ts.getRange('A2:G2').merge().setValue('Name of Adviser: ').setFontWeight('bold').setFontSize(14);
    }

    // 2. Setup Table Headers
    const headers = ['TIME', 'NO. OF MINUTES', 'MONDAY', 'TUESDAY', 'WEDNESDAY', 'THURSDAY', 'FRIDAY'];
    ts.getRange('A4:G4').setValues([headers]).setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle').setBorder(true,true,true,true,true,true);
    ts.setColumnWidth(1, 100); ts.setColumnWidth(2, 90); ts.setColumnWidths(3, 5, 130);

    // 3. Define Standard Timetable grid based on screenshots
    const grid = [
      { t: '7:00-7:30', m: 30, mon: 'Flag Raising\nCeremony', tue: 'GROUND PREPARATION/DAILY MORNING ROUTINE', wed: '', thu: '', fri: '', mergeSubj: true },
      { t: '7:30-8:30', m: 60, mon: '', tue: '', wed: '', thu: '', fri: '' },
      { t: '8:30-9:30', m: 60, mon: '', tue: '', wed: '', thu: '', fri: '' },
      { t: '9:30-9:45', m: 15, mon: 'RECESS/GROUP HANDWASHING/TOOTH BRUSHING', tue: '', wed: '', thu: '', fri: '', mergeSubj: true },
      { t: '9:45-10:45', m: 60, mon: '', tue: '', wed: '', thu: '', fri: '' },
      { t: '10:45-11:45', m: 60, mon: '', tue: '', wed: '', thu: '', fri: '' },
      { t: '11:45-1:00', m: 75, mon: 'LUNCH BREAK', tue: '', wed: '', thu: '', fri: '', mergeSubj: true },
      { t: '1:00-2:00', m: 60, mon: '', tue: '', wed: '', thu: '', fri: '' },
      { t: '1:00-2:30', m: 90, mon: '', tue: '', wed: '', thu: '', fri: '', shs: true },
      { t: '2:00-3:00', m: 60, mon: '', tue: '', wed: '', thu: '', fri: '' },
      { t: '2:30-3:30', m: 60, mon: '', tue: '', wed: '', thu: '', fri: '', shs: true },
      { t: '3:00-4:00', m: 60, mon: '', tue: '', wed: '', thu: '', fri: '' },
      { t: '3:30-4:00', m: 30, mon: '', tue: '', wed: '', thu: '', fri: '', shs: true },
    ];

    // Determine if it's SHS or Standard grid based on row data
    const isSHS = rows.some(r => {
       const grMatch = r[0].toString().match(/\b(11|12)\b/);
       return grMatch != null;
    });

    const activeGrid = grid.filter(g => {
       if (isSHS) {
         return !['1:00-2:00', '2:00-3:00', '3:00-4:00'].includes(g.t);
       } else {
         return !g.shs;
       }
    });

    // 4. Map the data into the grid
    let outputGrid = [];
    const days = ['MON', 'TUE', 'WED', 'THU', 'FRI'];
    const dIdx = { 'MON': 0, 'TUE': 1, 'WED': 2, 'THU': 3, 'FRI': 4 };

    // Function to find matching schedule for a given time slot and day
    const getSchedule = (timeStr, dayIndex) => {
       const [inStr, outStr] = timeStr.split('-');
       const sStart = parseTime(inStr.includes('AM') || inStr.includes('PM') ? inStr : inStr + (parseInt(inStr.split(':')[0]) < 7 || parseInt(inStr.split(':')[0]) === 12 ? ' PM' : ' AM'));

       const matched = rows.find(r => {
          if (r[2 + dayIndex] !== true) return false; // Days start at col 2
          const rStart = parseTime(r[7]); // TimeIn is col 7
          return Math.abs(rStart - sStart) < 15; // Match starts within 15 mins
       });

       if (!matched) return '';

       if (mode === 'Teacher') {
         return matched[1] + '\n' + matched[0]; // Subject + Section
       } else {
         return matched[1] + '\n(' + (matched[10] || 'Pending') + ')'; // Subject + Assigned Teacher
       }
    };

    activeGrid.forEach(g => {
       let m = g.mon, t = g.tue, w = g.wed, th = g.thu, f = g.fri;

       if (!g.mergeSubj) {
          m = getSchedule(g.t, 0) || m;
          t = getSchedule(g.t, 1) || t;
          w = getSchedule(g.t, 2) || w;
          th = getSchedule(g.t, 3) || th;
          f = getSchedule(g.t, 4) || f;
       }

       outputGrid.push([g.t, g.m, m, t, w, th, f]);
    });

    ts.getRange(5, 1, outputGrid.length, 7).setValues(outputGrid).setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true).setBorder(true,true,true,true,true,true);

    // Apply Merges for specific rows
    for (let i = 0; i < activeGrid.length; i++) {
       const r = 5 + i;
       if (activeGrid[i].t === '7:00-7:30') { ts.getRange(r, 4, 1, 4).merge().setBackground('#d9d9d9'); ts.getRange(r, 3).setBackground('#d9d9d9'); }
       if (activeGrid[i].t === '9:30-9:45') { ts.getRange(r, 3, 1, 5).merge().setBackground('#d9d9d9'); }
       if (activeGrid[i].t === '11:45-1:00') { ts.getRange(r, 3, 1, 5).merge().setBackground('#d9d9d9'); }
    }

    // 5. Build PDF URL
    SpreadsheetApp.flush();
    const url = ss.getUrl().replace(/edit$/, '') + 'export?exportFormat=pdf&format=pdf' +
      '&size=A4&portrait=false&fitw=true&sheetnames=false&printtitle=false&pagenumbers=false' +
      '&gridlines=false&fzr=false&gid=' + ts.getSheetId();

    const token = ScriptApp.getOAuthToken();
    const response = UrlFetchApp.fetch(url, { headers: { 'Authorization': 'Bearer ' + token } });

    const blob = response.getBlob().setName(targetName + '_Schedule.pdf');
    const file = DriveApp.createFile(blob);

    // Clean up
    ss.deleteSheet(ts);

    return file.getUrl();
  } catch (err) {
    if (ts) ss.deleteSheet(ts);
    return 'Error creating PDF: ' + err.toString();
  }
}
