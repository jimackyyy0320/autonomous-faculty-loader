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
  TERMS:          ['Term 1', 'Term 2', 'Term 3'],
  DASHBOARD:      'Teacher Dashboard',

  REPORT:         'Schedule Alignment',

  LUNCH_START:       705,  // 11:45 AM (mins since midnight)
  LUNCH_END:         780,  // 1:00 PM
  WEEKLY_WARN_HOURS: 28,
  WEEKLY_HARD_HOURS: 30,
  DAILY_PREFERRED_HOURS: 4.5,
  DAILY_WARN_HOURS:  4.5,
  DAILY_HARD_HOURS:  6,
  SLOT_STEP:         30,
  SCHOOL_START:      450,  // 7:30 AM
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
    .addItem('7️⃣ Build: All Sections Visualizer', 'buildAllSectionsVisualizer') // NEW VISUALIZER
    .addSeparator()
    .addItem('🚀 RUN AUTO-SCHEDULER', 'runAutoScheduler')
    .addItem('🔄 RECALCULATE SCHEDULE', 'runRecalculateSchedule')
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

  // 4.5 Interactive Learning: Real-time conflict feedback on manual edits
  if (CFG.TERMS.includes(sheetName) && row > 2 && (col === 3 || col >= 4 && col <= 10)) {
    checkRowConflicts(e.range.getSheet(), row);
    updateSubjectLoadingHours();
  }

  // 4.6 Subject Loading Live Updates
  if (sheetName === CFG.SUBJECT_LOAD && row > 2 && col <= 5) {
    updateSubjectLoadingHours();
  }

  // 1. Foundation Sweeps & Syncs
  if (sheetName === CFG.SECTION_ENROLL || sheetName === CFG.SUBJECT_LOAD || sheetName === CFG.TEACHER_ENROLL) {
    cleanOrphanedData();
    if (sheetName === CFG.SECTION_ENROLL) updateSectionDropdowns();

    if (sheetName === CFG.TEACHER_ENROLL) updateTeacherNameDropdowns();
  }

  // 2. Dashboards Live Filter (Both Teacher and Section)

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

  // 3.5 Term Tab Executable Auto-Fixer
  if (CFG.TERMS.includes(sheetName) && col === 11 && row > 2 && e.value === 'Fix Conflict') {
    applyTermFix(e);
  }
}

// ══════════════════════════════════════════════════════════════
//  SECTION 3: AUTO-SCHEDULER (The Super-Advanced Brain)
// ══════════════════════════════════════════════════════════════

function runRecalculateSchedule() {
  runAutoScheduler(true);
}

function runAutoScheduler(isReshuffle = false) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const subSheet = ss.getSheetByName(CFG.SUBJECT_LOAD);
  const enrollSheet = ss.getSheetByName(CFG.TEACHER_ENROLL);

  if (!subSheet || !enrollSheet) {
    return ss.toast('Please ensure your Data tabs are set up first.', '👋 Note', 5);
  }

  ss.toast('Running Strict Timeline-First Algorithm...', '🧠 Computing', 4);

  const getSafeData = (sheet, numCols) => {
    const lr = sheet.getLastRow();
    return lr < 3 ? [] : sheet.getRange(3, 1, lr - 2, numCols).getValues();
  };

  const demands = getSafeData(subSheet, 5).filter(r => r[0] && r[1] && r[2]); // Col E is Assigned Teacher
  const teachers = getSafeData(enrollSheet, 3).filter(r => r[1]);

  CFG.TERMS.forEach(term => {
    const ts = ss.getSheetByName(term);
    if (ts) {
      const lr = ts.getLastRow();
      if (lr >= 3) {
        ts.getRange(3, 1, lr - 2, 13).clearContent();
        ts.getRange(3, 4, lr - 2, 5).insertCheckboxes();
        const actionRule = SpreadsheetApp.newDataValidation().requireValueInList(['—', 'Fix Conflict'], true).build();
        ts.getRange(3, 11, lr - 2, 1).setDataValidation(actionRule);
      }
    }
  });

  let unmappedLog = [];
  const props = PropertiesService.getDocumentProperties();

  CFG.TERMS.forEach(term => {
    const termDemands = demands.filter(d => d[0] === term);
    if (!termDemands.length) return;

    const tBooked = {};
    const tBookedMins = {};
    const tBookedSubjects = {}; // Tracks { day: Set(subject_section) } for non-consecutive rules
    const cBooked = {};
    const outputRows = [];

    // Initialize tracking structures
    const sectionDemands = {};
    termDemands.forEach(d => {
      const sec = d[1].toString().trim();
      const subj = d[2].toString().trim();
      const hrs = parseFloat(d[3]) || 0;

      const learnedTeacher = props.getProperty('LEARNED_' + subj + '|' + sec);
      const teacher = learnedTeacher || (d[4] ? d[4].toString().trim() : '') || '⚠️ Unassigned';

      if (!tBooked[teacher] && teacher !== 'Unavailable Teacher' && teacher !== '⚠️ Unassigned') {
          tBooked[teacher] = { 1:[], 2:[], 3:[], 4:[], 5:[], sectionsMet: { 1:new Set(), 2:new Set(), 3:new Set(), 4:new Set(), 5:new Set() } };
          tBookedMins[teacher] = { 1:0, 2:0, 3:0, 4:0, 5:0 };
          tBookedSubjects[teacher] = { 1:[], 2:[], 3:[], 4:[], 5:[] };
      }
      if (!cBooked[sec]) cBooked[sec] = { 1:[], 2:[], 3:[], 4:[], 5:[] };

      if (!sectionDemands[sec]) sectionDemands[sec] = [];

      // Determine Duration Logic
      const isPhilGov = subj.toLowerCase().includes('phil gov') || subj.toLowerCase().includes('philippine politics');
      const gradeMatch = sec.match(/\b(11|12|7|8|9|10)\b/);
      const grade = gradeMatch ? parseInt(gradeMatch[1], 10) : 7;

      let durationMins = 60; // Minimum 1 hour unless Homeroom
      if (isPhilGov) durationMins = 90;
      else if (hrs % 1 === 0.5 && !isHomeroom) durationMins = 90; // Fractional weekly hours end in .5 -> 90 mins

      if (isHomeroom && hrs === 0.5) durationMins = 30; // Homeroom exception

      const isAral = subj.match(/\bARAL\b/i) && !subj.toLowerCase().includes('panlipunan');
      if (isAral) durationMins = 60; // ARAL exactly 1 hour

      const isHomeroom = subj.toLowerCase().includes('homeroom');

      sectionDemands[sec].push({
         subject: subj,
         teacher: teacher,
         hoursLeft: hrs,
         originalHours: hrs,
         durationMins: durationMins,
         isAral: isAral,
         isHomeroom: isHomeroom,
         grade: grade,
         slotsAcquired: []
      });
    });

    const sectionKeys = Object.keys(sectionDemands);
    if (isReshuffle) sectionKeys.sort(() => Math.random() - 0.5);

    // DISTRIBUTE UNASSIGNED SUBJECTS (Cross-Specialization Load Balancing)
    let teacherWeeklyHours = {};
    teachers.forEach(t => { if (t[1]) teacherWeeklyHours[t[1]] = 0; });

    // First pass: Calculate pre-assigned hours
    sectionKeys.forEach(sec => {
        sectionDemands[sec].forEach(d => {
            if (d.teacher !== '⚠️ Unassigned' && d.teacher !== 'Unavailable Teacher') {
                if (teacherWeeklyHours[d.teacher] !== undefined) {
                    teacherWeeklyHours[d.teacher] += d.originalHours;
                }
            }
        });
    });

    // Second pass: Distribute unassigned
    sectionKeys.forEach(sec => {
        sectionDemands[sec].forEach(d => {
            if (d.teacher === '⚠️ Unassigned') {
                let bestT = null;
                let lowestScore = Infinity;

                let candidateTeachers = teachers.map(t => t[1]).filter(t => t !== 'Unavailable Teacher' && t !== '⚠️ Unassigned');
                candidateTeachers.sort(() => Math.random() - 0.5);

                candidateTeachers.forEach(tName => {
                    let hrs = teacherWeeklyHours[tName] || 0;

                    let isSpecMatch = false;
                    let tObj = teachers.find(t => t[1] === tName);
                    if (tObj && tObj[2]) {
                        let spec = tObj[2].toString().toLowerCase();
                        let subjLower = d.subject.toLowerCase();
                        if (subjLower.includes(spec) || spec.includes(subjLower)) isSpecMatch = true;
                    }

                    // Score = current hours + penalty for non-specialist
                    let score = hrs;
                    if (!isSpecMatch) score += 5;

                    if (score < lowestScore) {
                        lowestScore = score;
                        bestT = tName;
                    }
                });

                if (bestT) {
                    d.teacher = bestT;
                    d.isSub = true; // flag as distributed substitute
                    teacherWeeklyHours[bestT] += d.originalHours;

                    if (!tBooked[bestT]) {
                        tBooked[bestT] = { 1:[], 2:[], 3:[], 4:[], 5:[], sectionsMet: { 1:new Set(), 2:new Set(), 3:new Set(), 4:new Set(), 5:new Set() } };
                        tBookedMins[bestT] = { 1:0, 2:0, 3:0, 4:0, 5:0 };
                        tBookedSubjects[bestT] = { 1:[], 2:[], 3:[], 4:[], 5:[] };
                    }
                }
            }
        });
    });

    // TIMELINE-FIRST SCHEDULING (Dense Packing)
    // We iterate over every Section, then Day, then Time. We strictly map gaps.

    // Sort sections so reshuffle affects evaluation order

    sectionKeys.forEach(section => {
       const demands = sectionDemands[section];
       const gradeMatch = section.match(/\b(11|12|7|8|9|10)\b/);
       const grade = gradeMatch ? parseInt(gradeMatch[1], 10) : 7;

       // Pre-map Homeroom strictly
       demands.forEach(d => {
         if (d.isHomeroom && d.hoursLeft > 0) {
            const day = 1; // Monday
            let sStart = 450; // 7:30 AM
            if (grade >= 11) sStart = 900; // 3:00 PM
            const sEnd = sStart + 60;

            cBooked[section][day].push({s: sStart, e: sEnd, subject: d.subject});
            if (tBooked[d.teacher]) {
               tBooked[d.teacher][day].push({s: sStart, e: sEnd});
               tBookedMins[d.teacher][day] += 60;
               tBooked[d.teacher].sectionsMet[day].add(section);
            }
            d.slotsAcquired.push({ day, in: formatMinsToTime(sStart), out: formatMinsToTime(sEnd), s: sStart, e: sEnd, assignedT: d.teacher });
            d.hoursLeft -= 1;
         }
       });

       let days = [1,2,3,4,5];
       if (isReshuffle) days.sort(() => Math.random() - 0.5);

       days.forEach(day => {
          let time = 450; // Start at 7:30 AM

          while (time < CFG.SCHOOL_END) { // Until 4:30 PM
             // Skip pre-booked slots (like homeroom or manually inserted things)
             const existingBlock = cBooked[section][day].find(b => time >= b.s && time < b.e);
             if (existingBlock) {
                 time = existingBlock.e;
                 continue;
             }

             // Handle Recess (9:30 AM - 9:45 AM = 570-585)
             if (time === 570) { time = 585; continue; }
             // Handle Lunch (11:45 AM - 1:00 PM = 705-780)
             if (time === 705) { time = 780; continue; }

             // Find the best subject to fit in this time slot to guarantee dense packing
             // We want to avoid skipping. If no subject perfectly fits, we MUST map an "Unavailable Teacher"
             // to the best matching subject to maintain dense packing.

             let validCandidates = demands.filter(d => d.hoursLeft > 0);
             if (validCandidates.length === 0) break; // Finished schedule for section

             // Filter by Hard Constraints
             let bestSubj = null;
             let bestScore = -Infinity;
             let forceUnavailable = false;

             validCandidates.forEach(d => {
                let score = 100;
                let tBlocked = false;
                let sBlocked = false; // Section blocked by logical constraint (not time conflict)

                // 1. Durations & Boundaries
                let blockDur = d.durationMins;
                if (d.hoursLeft * 60 < blockDur) {
                    blockDur = d.hoursLeft * 60; // Truncate last block
                }
                // Minimum 1 hour rule unless Homeroom
                if (!d.isHomeroom && blockDur < 60) blockDur = 60;
                const endTime = time + blockDur;

                // Crosses Recess or Lunch boundary? Invalid.
                if (time < 570 && endTime > 570) sBlocked = true;
                if (time < 705 && endTime > 705) sBlocked = true;
                if (endTime > CFG.SCHOOL_END) sBlocked = true;

                // 2. ARAL Placement (Must be exactly 3:00 PM - 4:00 PM)
                if (d.isAral && time !== 900) sBlocked = true;
                if (!d.isAral && time === 900) {
                   // If another subject tries to take 3-4 PM, check if ARAL still needs it.
                   const aralNeeds = demands.some(a => a.isAral && a.hoursLeft > 0);
                   if (aralNeeds) score -= 1000; // Leave 3PM open for ARAL
                }

                // 3. G7-10 Meeting Frequency
                const alreadyMetToday = cBooked[section][day].some(b => b.subject === d.subject);
                if (grade <= 10 && alreadyMetToday) sBlocked = true;

                // 4. G11-12 Consecutive Check
                if (grade >= 11 && alreadyMetToday) {
                   // Allowed to meet 2x, but NOT consecutively.
                   const prevBlock = cBooked[section][day].find(b => b.subject === d.subject);
                   if (prevBlock && (prevBlock.e === time || prevBlock.s === endTime)) {
                       sBlocked = true; // Back to back
                   }
                }

                if (sBlocked) return; // Cannot evaluate this subject for this slot

                // Evaluate Teacher Constraints
                const t = d.teacher;
                if (t === '⚠️ Unassigned' || t === 'Unavailable Teacher') tBlocked = true;
                else if (tBooked[t]) {
                   // Time conflict
                   if (tBooked[t][day].some(b => time < b.e && endTime > b.s)) tBlocked = true;

                   // G7-10 Advisory Frequency for Teacher (Fix: Check if this teacher is the section's homeroom adviser)
                   // The homeroom teacher for a section is mapped early. Let's determine if 't' is the homeroom teacher.
                   const isAdviser = demands.some(x => x.isHomeroom && x.teacher === t);
                   if (grade <= 10 && isAdviser && tBooked[t].sectionsMet[day].has(section)) tBlocked = true;

                   // 6-Hour Hard Limit
                   if (tBookedMins[t][day] + blockDur > CFG.DAILY_HARD_HOURS * 60) tBlocked = true;

                   // Soft Constraints (Heuristics) if teacher is free
                   if (!tBlocked) {
                       // Prefer <= 4.5 hours (4.5h config)
                       if (tBookedMins[t][day] + blockDur > CFG.DAILY_PREFERRED_HOURS * 60) score -= 150;

                       // Teacher Prep Time: Penalize back-to-back
                       const hasBackToBack = tBooked[t][day].some(b => b.e === time || b.s === endTime);
                       if (hasBackToBack) score -= 20;

                       // Perfect Block: Was it scheduled at this time on previous days?
                       let sameTimeCount = 0;
                       d.slotsAcquired.forEach(sa => {
                          if (sa.s === time) sameTimeCount++;
                       });
                       score += sameTimeCount * 15;
                   }
                }

                if (tBlocked) score -= 1000; // FIX: Heavily penalize blocked teachers so they never win against valid soft-penalties

                // Sort weighting
                score += d.hoursLeft * 2; // Prioritize heavier subjects

                if (score > bestScore) {
                   bestScore = score;
                   bestSubj = d;
                   forceUnavailable = tBlocked;
                }
             });

             if (bestSubj) {
                // Book it!
                const bDur = Math.min(bestSubj.durationMins, bestSubj.hoursLeft * 60);
                const eTime = time + bDur;

                const finalTeacher = forceUnavailable ? 'Unavailable Teacher' : bestSubj.teacher;

                cBooked[section][day].push({s: time, e: eTime, subject: bestSubj.subject});

                if (finalTeacher !== 'Unavailable Teacher' && tBooked[finalTeacher]) {
                    tBooked[finalTeacher][day].push({s: time, e: eTime});
                    tBookedMins[finalTeacher][day] += bDur;
                    tBooked[finalTeacher].sectionsMet[day].add(section);
                }

                bestSubj.slotsAcquired.push({
                   day: day, in: formatMinsToTime(time), out: formatMinsToTime(eTime),
                   s: time, e: eTime, assignedT: finalTeacher
                });

                bestSubj.hoursLeft -= (bDur / 60);
                time = eTime; // Advance time
             } else {
                // No subject could legally fit here (e.g. ARAL constraints, etc.)
                // To maintain dense packing, we are forced to skip the time slot. This is a severe bottleneck.
                time += 30; // Advance time to try next slot
             }
          }
       });
    });

    // Process output rows
    sectionKeys.forEach(section => {
       const demands = sectionDemands[section];
       demands.forEach(d => {
          if (d.hoursLeft > 0) {
            unmappedLog.push(`[${term}] ${d.subject} (${section}) — ${d.hoursLeft}h unmapped. No physical slots left in week.`);
          }

          const grouped = {};
          d.slotsAcquired.forEach(sa => {
            const key = sa.in + '|' + sa.out + '|' + sa.assignedT;
            if (!grouped[key]) grouped[key] = { m:false, t:false, w:false, th:false, f:false, in: sa.in, out: sa.out, finalTeacher: sa.assignedT };
            if (sa.day === 1) grouped[key].m = true;
            if (sa.day === 2) grouped[key].t = true;
            if (sa.day === 3) grouped[key].w = true;
            if (sa.day === 4) grouped[key].th = true;
            if (sa.day === 5) grouped[key].f = true;
          });

          // Teacher Suggestions Based on Major
          let suggestions = [];
          const subjLower = d.subject.toLowerCase();
          teachers.forEach(t => {
            const spec = t[2].toString().toLowerCase();
            if (spec && (subjLower.includes(spec) || spec.includes(subjLower))) {
              suggestions.push(t[1]);
            }
          });
          if (suggestions.length === 0) {
            teachers.forEach(t => {
              const spec = t[2].toString().toLowerCase();
              const words = subjLower.split(' ').filter(w => w.length > 3);
              if (words.some(w => spec.includes(w))) suggestions.push(t[1]);
            });
          }
          const suggStr = suggestions.length > 0 ? suggestions.slice(0, 3).join(', ') : 'Any Teacher';

          Object.values(grouped).forEach(g => {
            let refinedSugg = suggStr;
            let warn = '';

            if (g.finalTeacher === 'Unavailable Teacher') {
                warn = '🔴 Teacher Conflict';
                // Find strictly free teachers
                const startMins = parseTime(g.in);
                const endMins = parseTime(g.out);
                const activeDays = [];
                if (g.m) activeDays.push(1); if (g.t) activeDays.push(2); if (g.w) activeDays.push(3); if (g.th) activeDays.push(4); if (g.f) activeDays.push(5);

                let trulyFree = [];
                suggestions.forEach(tName => {
                   let isBlocked = false;
                   if (!tBooked[tName]) return;
                   activeDays.forEach(day => {
                      if (tBooked[tName][day].some(b => startMins < b.e && endMins > b.s)) isBlocked = true;
                      // Enforce 6h limit for suggestion
                      if (tBookedMins[tName][day] + (endMins - startMins) > CFG.DAILY_HARD_HOURS * 60) isBlocked = true;
                   });
                   if (!isBlocked) trulyFree.push(tName);
                });
                if (trulyFree.length > 0) refinedSugg = '★ ' + trulyFree.slice(0, 3).join(', ');
            } else {
                // Check if they exceeded 4.5h on any of these days
                const dayMap = {m:1, t:2, w:3, th:4, f:5};
                let softLimitHit = false;
                ['m','t','w','th','f'].forEach(dayKey => {
                    if (g[dayKey]) {
                        let dIdx = dayMap[dayKey];
                        if (tBookedMins[g.finalTeacher] && tBookedMins[g.finalTeacher][dIdx] > CFG.DAILY_PREFERRED_HOURS * 60) {
                            softLimitHit = true;
                        }
                    }
                });

                if (softLimitHit) {
                    warn = warn ? warn + ', 🟠 >4.5h Soft Limit' : '🟠 >4.5h Soft Limit';
                }

                if (d.isSub) {
                    warn = warn ? warn + ', 🟢 Distributed Sub' : '🟢 Distributed Sub';
                }
            }

            outputRows.push([section, d.subject, g.finalTeacher, g.m, g.t, g.w, g.th, g.f, g.in, g.out, '—', warn, refinedSugg]);
          });
       });
    });

    const ts = ss.getSheetByName(term);
    if (ts && outputRows.length > 0) {
      ts.getRange(3, 1, outputRows.length, 13).setValues(outputRows);
      ts.getRange(3, 1, outputRows.length, 13).setVerticalAlignment('middle');
      ts.getRange(3, 9, outputRows.length, 2).setHorizontalAlignment('center');
      ts.getRange(3, 12, outputRows.length, 1).setFontColor(C.warn).setFontStyle('italic');
    }
  });

  if (unmappedLog.length > 0) {
    const ui = SpreadsheetApp.getUi();
    ui.alert(
      '⚠️ Timetable Adjustments Required',
      'Some sections have physical unmapped hours due to extreme bottlenecks:\n\n' + unmappedLog.join('\n\n') + '\n\nPlease review the Term tabs.',
      ui.ButtonSet.OK
    );

    const rep = ss.getSheetByName(CFG.REPORT);
    if (rep) {
        let lr = Math.max(3, rep.getLastRow() + 1);
        let out = [];
        unmappedLog.forEach(log => {
            // log format: [Term 1] Science (Grade 7) — 1h unmapped. No physical slots left in week.
            const match = log.match(/\[(.*?)\] (.*?) \((.*?)\) — (.*?h) unmapped/);
            if (match) {
                out.push([match[1], '🔴 Unmapped Hours', 'ALL', match[3], 'N/A', 'N/A', match[2], `Needs ${match[4]}`, '—']);
            }
        });
        if (out.length > 0) {
            rep.getRange(lr, 1, out.length, 9).setValues(out).setBackground(C.errorBg).setFontColor(C.error);
        }
    }
  } else {
    ss.toast('Schedules are completely mapped out and optimized!', '✅ All Set', 6);
  }

  if (typeof updateSubjectLoadingHours === 'function') updateSubjectLoadingHours();
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

function _buildTermEntries(src, termName) {
  const last = src.getLastRow();
  if (last < 3) return [];
  const rows = src.getRange(3, 1, last - 2, 10).getValues();

  return rows.flatMap((row, idx) => {
    const gradeLevel = row[0], subject = row[1], teacher = row[2];
    const timeIn = row[8], timeOut = row[9];
    if (!teacher || !timeIn || !timeOut) return [];

    const start = parseTime(timeIn);
    const end = parseTime(timeOut);
    if (start >= end) return [];

    const days = ['MON','TUE','WED','THU','FRI'];
    return days.map((day, di) => ({ day, di }))
      .filter(({ di }) => row[3 + di] === true)
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

function applyTermFix(e) {
  const sheet = e.range.getSheet();
  const row = e.range.getRow();

  const suggestions = sheet.getRange(row, 13).getValue();
  let newTeacher = '';

  if (suggestions && suggestions.startsWith('★ ')) {
      newTeacher = suggestions.replace('★ ', '').split(',')[0].trim();
  } else if (suggestions && suggestions !== 'Any Teacher') {
      newTeacher = suggestions.split(',')[0].trim();
  }

  if (newTeacher) {
      sheet.getRange(row, 3).setValue(newTeacher);
      e.range.setValue('✔ Fixed').setBackground(C.okBg).setFontColor(C.ok).clearDataValidations();
      e.source.toast(`Reassigned class to ${newTeacher}.`, '✅ Auto-Fixed', 4);

      if (typeof checkRowConflicts === 'function') checkRowConflicts(sheet, row);
      if (typeof updateSubjectLoadingHours === 'function') updateSubjectLoadingHours();
  } else {
      e.range.setValue('—');
      e.source.toast(`No valid alternative teacher found for auto-fix. Please fix manually.`, '⚠️ Alert', 5);
  }
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
//  SECTION 4.5: INTERACTIVE LEARNING & REAL-TIME CONFLICTS
// ══════════════════════════════════════════════════════════════
function checkRowConflicts(sheet, rowNum) {
  const termName = sheet.getName();
  const data = sheet.getRange(3, 1, Math.max(1, sheet.getLastRow() - 2), 10).getValues();

  const targetRow = data[rowNum - 3];
  if (!targetRow) return;

  const section = targetRow[0];
  const subject = targetRow[1];
  const teacher = targetRow[2];
  const tIn = parseTime(targetRow[8]);
  const tOut = parseTime(targetRow[9]);

  // AI Learning Module: Memorize manual overrides
  if (section && subject && teacher) {
    const props = PropertiesService.getDocumentProperties();
    props.setProperty('LEARNED_' + subject + '|' + section, teacher);
  }

  if (!teacher || tIn >= tOut || teacher === 'Unavailable Teacher') {
    sheet.getRange(rowNum, 12).setValue('');
    return;
  }

  const activeDays = [3, 4, 5, 6, 7].filter(d => targetRow[d] === true);
  if (activeDays.length === 0) return;

  let overlaps = [];
  let dailyHrs = {3:0, 4:0, 5:0, 6:0, 7:0};

  data.forEach((r, idx) => {
    if (idx === rowNum - 3) return; // skip self
    const rStart = parseTime(r[8]);
    const rEnd = parseTime(r[9]);
    if (rStart >= rEnd) return;

    // Check overlaps
    const sameTeacher = r[2] === teacher;
    const sameSection = r[0] === section && section !== '';
    const sameSubject = sameSection && r[1] === subject;

    if (sameTeacher || sameSection) {
      activeDays.forEach(d => {
        if (r[d] === true) {
          if (tIn < rEnd && tOut > rStart) {
            overlaps.push(sameTeacher ? '🔴 Teacher Overlap' : '🔴 Section Overlap');
          }
          // Back-to-back constraint for same subject (G11-12) or any (G7-10)
          if (sameSubject) {
            const gradeMatch = section.match(/\b(11|12|7|8|9|10)\b/);
            const grade = gradeMatch ? parseInt(gradeMatch[1], 10) : 7;
            if (grade <= 10) overlaps.push('🔴 JHS: Met twice in a day');
            else if (tIn === rEnd || tOut === rStart) overlaps.push('🔴 SHS: Back-to-back same subject');
          }
        }
      });
    }

    // Accumulate teacher hours
    if (sameTeacher) {
      activeDays.forEach(d => {
        if (r[d] === true) dailyHrs[d] += (rEnd - rStart) / 60;
      });
    }
  });

  // Add current row duration to daily hrs
  const dur = (tOut - tIn) / 60;
  activeDays.forEach(d => dailyHrs[d] += dur);

  let warnings = [];
  if (overlaps.length > 0) warnings.push(overlaps[0]);

  activeDays.forEach(d => {
    if (dailyHrs[d] > CFG.DAILY_HARD_HOURS) warnings.push('🔴 Over 6h limit');
    else if (dailyHrs[d] > CFG.DAILY_PREFERRED_HOURS) warnings.push('🟠 Over 4.5h limit');
  });

  const warnCell = sheet.getRange(rowNum, 12);
  if (warnings.length > 0) {
    warnings = [...new Set(warnings)]; // Dedup
    const hasError = warnings.some(w => w.includes('🔴'));
    warnCell.setValue(warnings.join(', ')).setFontColor(hasError ? C.error : C.warn).setFontStyle('normal').setFontWeight('bold');
  } else {
    warnCell.setValue('✅ Valid').setFontColor(C.ok).setFontStyle('italic').setFontWeight('normal');
  }
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

  const dayMap = { MON: 3, TUE: 4, WED: 5, THU: 6, FRI: 7 };
  const lastRow = src.getLastRow();
  if (lastRow < 3) return;
  const rows = src.getRange(3, 1, lastRow - 2, 13).getValues();

  let output = [];
  let totalMins = 0;

  rows.forEach(row => {
    if (row[2] !== teacher) return;

    const start = parseTime(row[8]);
    const end   = parseTime(row[9]);
    const dur   = calcEffectiveDuration(start, end);

    if (day === 'ALL WEEK') {
      let daysOn = 0;
      for (let d = 3; d <= 7; d++) { if (row[d] === true) daysOn++; }
      totalMins += dur * daysOn;
      output.push([row[1], row[0], row[8], row[9], row[3], row[4], row[5], row[6], row[7]]);
    } else {
      if (row[dayMap[day]] !== true) return;
      totalMins += dur;
      output.push([row[1], row[0], row[8], row[9], '', '', '', '', '']);
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

  const dayMap = { MON: 3, TUE: 4, WED: 5, THU: 6, FRI: 7 };
  const lastRow = src.getLastRow();
  if (lastRow < 3) return;
  const rows = src.getRange(3, 1, lastRow - 2, 10).getValues();

  let output = [];
  let totalMins = 0;

  rows.forEach(row => {
    if (row[0] !== section) return; // Match Grade Level/Section instead of Teacher

    const start = parseTime(row[8]);
    const end   = parseTime(row[9]);
    const dur   = calcEffectiveDuration(start, end);

    if (day === 'ALL WEEK') {
      let daysOn = 0;
      for (let d = 3; d <= 7; d++) { if (row[d] === true) daysOn++; }
      totalMins += dur * daysOn;
      // Note the column swap: Subject, Teacher, TimeIn...
      output.push([row[1], row[2], row[8], row[9], row[3], row[4], row[5], row[6], row[7]]);
    } else {
      if (row[dayMap[day]] !== true) return;
      totalMins += dur;
      output.push([row[1], row[2], row[8], row[9], '', '', '', '', '']);
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
  const teachEnrollSheet = ss.getSheetByName(CFG.TEACHER_ENROLL);

  if (!secSheet || !subSheet || !teachEnrollSheet) return;

  const validSections = new Set(secSheet.getRange('B3:B').getValues().flat().filter(String));
  const validTeachers = new Set(teachEnrollSheet.getRange('B3:B').getValues().flat().filter(String));
  validTeachers.add('Unavailable Teacher');
  validTeachers.add('⚠️ Unassigned');

  const subLastRow = Math.max(3, subSheet.getLastRow());
  const subRange = subSheet.getRange('B3:E' + subLastRow); // Get Section and Teacher
  const subValues = subRange.getValues();
  let subChanged = false;

  subValues.forEach(row => {
    if (row[0] && !validSections.has(row[0])) { row[0] = ''; subChanged = true; }
    if (row[3] && !validTeachers.has(row[3])) { row[3] = ''; subChanged = true; } // row[3] is Column E
  });
  if (subChanged) subRange.setValues(subValues);
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
  const subSheet = ss.getSheetByName(CFG.SUBJECT_LOAD);
  const dashSheet   = ss.getSheetByName(CFG.DASHBOARD);

  if (!enrollSheet) return;

  const tVals = enrollSheet.getRange('B3:B1000').getValues().flat().filter(String);
  tVals.push('Unavailable Teacher');
  tVals.push('⚠️ Unassigned');
  const rule = SpreadsheetApp.newDataValidation().requireValueInList(tVals, true).build();

  if (subSheet) subSheet.getRange('E3:E1000').setDataValidation(rule);
  if (dashSheet) {
    dashSheet.getRange('B2').setDataValidation(rule);
    dashSheet.getRange('L2').setDataValidation(rule);
  }
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
  subSheet.getRange('A1:G1').merge().setValue('📚 SUBJECT LOADING (CURRICULUM DEMAND & ASSIGNMENT)').setFontWeight('bold').setFontSize(12).setBackground(C.navyDark).setFontColor(C.white).setHorizontalAlignment('center').setVerticalAlignment('middle');
  subSheet.getRange('A2:G2').setValues([['Term', 'Section', 'Subject', 'Weekly Class Hours', 'Assigned Teacher', 'Hours Loaded', 'Balance']]).setFontWeight('bold').setFontSize(10).setBackground(C.teal).setFontColor(C.white).setHorizontalAlignment('center').setVerticalAlignment('middle');
  subSheet.setColumnWidth(1, 120); subSheet.setColumnWidth(2, 200); subSheet.setColumnWidth(3, 250); subSheet.setColumnWidth(4, 150); subSheet.setColumnWidth(5, 200); subSheet.setColumnWidth(6, 120); subSheet.setColumnWidth(7, 120);
  subSheet.getRange('A3:G1000').setFontColor(C.body).setVerticalAlignment('middle');
  subSheet.getRange('D3:D1000').setNumberFormat('0.0').setHorizontalAlignment('center');
  subSheet.getRange('F3:G1000').setNumberFormat('0.0').setHorizontalAlignment('center');
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

  updateTeacherNameDropdowns();
  ss.toast('Teacher Enrollment initialized.', '✅ Phase 2 Complete', 5);
}

function buildPhase3TermTabs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  CFG.TERMS.forEach((term, index) => {
    let sheet = ss.getSheetByName(term) || ss.insertSheet(term, 4 + index);
    sheet.clear();
    sheet.getRange('A1:J1').merge().setValue(`📅 ${term.toUpperCase()} MASTER SCHEDULE`).setFontWeight('bold').setFontSize(12).setBackground(C.navyDark).setFontColor(C.white).setHorizontalAlignment('center').setVerticalAlignment('middle');
    sheet.setRowHeight(1, 40);

    const headers = ['GRADE Level', 'Subject', 'Teacher', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Time In', 'Time Out', 'Action', 'Warnings', 'Suggested Teachers'];
    sheet.getRange('A2:M2').setValues([headers]).setFontWeight('bold').setFontSize(10).setBackground(C.teal).setFontColor(C.white).setHorizontalAlignment('center').setVerticalAlignment('middle');
    sheet.setRowHeight(2, 30);

    sheet.setColumnWidth(1, 150); sheet.setColumnWidth(2, 250); sheet.setColumnWidth(3, 200);
    sheet.setColumnWidths(4, 5, 80); sheet.setColumnWidths(9, 2, 110);

    sheet.getRange('A3:J1000').setFontColor(C.body).setVerticalAlignment('middle');

    const cbRule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
    sheet.getRange('D3:H1000').setDataValidation(cbRule).setHorizontalAlignment('center');

    sheet.getRange('I3:J1000').setHorizontalAlignment('center').setNumberFormat('h:mm AM/PM');
    sheet.getRange('A3:A1000').setHorizontalAlignment('center');
    sheet.getRange('B3:C1000').setHorizontalAlignment('left');
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




function updateSubjectLoadingHours() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const subSheet = ss.getSheetByName(CFG.SUBJECT_LOAD);
  if (!subSheet) return;

  const lr = Math.max(3, subSheet.getLastRow());
  const demands = subSheet.getRange(3, 1, lr - 2, 4).getValues(); // Read up to Col D (Weekly Class Hours)

  // Aggregate all loaded hours from Term tabs
  const loadedMap = {};

  CFG.TERMS.forEach(term => {
     const ts = ss.getSheetByName(term);
     if (!ts) return;
     const tlr = ts.getLastRow();
     if (tlr < 3) return;
     const rows = ts.getRange(3, 1, tlr - 2, 10).getValues(); // 0:Sec, 1:Subj, 2:Teach, 3-7:Days, 8:In, 9:Out

     rows.forEach(r => {
        const sec = r[0].toString().trim();
        const sub = r[1].toString().trim();
        if (!sec || !sub) return;

        const start = parseTime(r[8]);
        const end = parseTime(r[9]);
        if (start >= end) return;

        let daysOn = 0;
        for (let d = 3; d <= 7; d++) { if (r[d] === true) daysOn++; }

        const key = term + '|' + sec + '|' + sub;
        if (!loadedMap[key]) loadedMap[key] = 0;

        // Add duration mapped across all active days for this slot
        loadedMap[key] += ((end - start) / 60) * daysOn;
     });
  });

  const output = [];
  let bgColors = [];

  demands.forEach(d => {
     const key = d[0] + '|' + d[1] + '|' + d[2];
     const req = parseFloat(d[3]) || 0;
     const loaded = loadedMap[key] || 0;
     const balance = req - loaded;

     if (!d[0] || !d[1] || !d[2]) {
         output.push(['', '']);
         bgColors.push([C.white, C.white]);
     } else {
         output.push([loaded, balance]);
         // Status coloring
         if (loaded < req) bgColors.push([C.warnBg, C.warnBg]);
         else if (loaded > req) bgColors.push([C.errorBg, C.errorBg]);
         else bgColors.push([C.okBg, C.okBg]);
     }
  });

  if (output.length > 0) {
      subSheet.getRange(3, 6, output.length, 2).setValues(output).setBackgrounds(bgColors);
  }
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
//  SECTION 10: MASTER SECTION VISUALIZER
// ══════════════════════════════════════════════════════════════

function buildAllSectionsVisualizer() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const visName = '📚 All Sections Schedule';
  let vis = ss.getSheetByName(visName) || ss.insertSheet(visName, 5);
  vis.clear();

  vis.getRange('A1:G1').merge().setValue('📚 MASTER SECTIONS TIMETABLE VISUALIZER').setFontWeight('bold').setFontSize(14).setBackground(C.navyDark).setFontColor(C.white).setHorizontalAlignment('center').setVerticalAlignment('middle');
  vis.setRowHeight(1, 46);
  vis.setColumnWidth(1, 100); vis.setColumnWidth(2, 90); vis.setColumnWidths(3, 5, 140);

  let allRows = [];
  CFG.TERMS.forEach(term => {
     const ts = ss.getSheetByName(term);
     if (!ts) return;
     const lr = ts.getLastRow();
     if (lr >= 3) {
        const rows = ts.getRange(3, 1, lr - 2, 11).getValues();
        allRows = allRows.concat(rows);
     }
  });

  if (allRows.length === 0) {
     return ss.toast('No schedules found in Term tabs to visualize.', '⚠️ Empty', 4);
  }

  const sections = [...new Set(allRows.map(r => r[0] ? r[0].toString().trim() : '').filter(String))];

  sections.sort((a, b) => {
     const matchA = a.match(/\d+/); const matchB = b.match(/\d+/);
     const numA = matchA ? parseInt(matchA[0], 10) : 0;
     const numB = matchB ? parseInt(matchB[0], 10) : 0;
     if (numA !== numB) return numA - numB;
     return a.localeCompare(b);
  });

  let currentRow = 3;

  sections.forEach(section => {
      const secRows = allRows.filter(r => r[0] && r[0].toString().trim() === section);
      if (secRows.length === 0) return;

      const isSHS = section.match(/\b(11|12)\b/) != null;

      vis.getRange(currentRow, 1, 1, 7).merge().setValue('Cohort / Section: ' + section).setFontWeight('bold').setFontSize(12).setBackground(C.teal).setFontColor(C.white).setHorizontalAlignment('left').setVerticalAlignment('middle');
      currentRow++;

      const headers = ['TIME', 'MINS', 'MONDAY', 'TUESDAY', 'WEDNESDAY', 'THURSDAY', 'FRIDAY'];
      vis.getRange(currentRow, 1, 1, 7).setValues([headers]).setFontWeight('bold').setBackground(C.navyLight).setHorizontalAlignment('center').setVerticalAlignment('middle').setBorder(true,true,true,true,true,true);
      currentRow++;

      const grid = [
        { t: '7:00-7:30', m: 30, mon: 'Flag Raising\nCeremony', tue: 'GROUND PREPARATION/DAILY MORNING ROUTINE', wed: '', thu: '', fri: '', mergeSubj: true },
        { t: '7:30-8:30', m: 60, mon: '', tue: '', wed: '', thu: '', fri: '' },
        { t: '8:30-9:30', m: 60, mon: '', tue: '', wed: '', thu: '', fri: '' },
        { t: '9:30-9:45', m: 15, mon: 'RECESS/GROUP HANDWASHING', tue: '', wed: '', thu: '', fri: '', mergeSubj: true },
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

      const activeGrid = grid.filter(g => isSHS ? !['1:00-2:00', '2:00-3:00', '3:00-4:00'].includes(g.t) : !g.shs);
      let outputGrid = [];

      const getSchedule = (timeStr, dayIndex) => {
         const [inStr, outStr] = timeStr.split('-');
         const sStart = parseTime(inStr.includes('AM') || inStr.includes('PM') ? inStr : inStr + (parseInt(inStr.split(':')[0]) < 7 || parseInt(inStr.split(':')[0]) === 12 ? ' PM' : ' AM'));
         const matched = secRows.find(r => r[3 + dayIndex] === true && Math.abs(parseTime(r[8]) - sStart) < 15);
         return matched ? matched[1] + '\n(' + matched[2] + ')' : '';
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

      vis.getRange(currentRow, 1, outputGrid.length, 7).setValues(outputGrid).setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true).setBorder(true,true,true,true,true,true);

      for (let i = 0; i < activeGrid.length; i++) {
         const r = currentRow + i;
         if (activeGrid[i].t === '7:00-7:30') { vis.getRange(r, 4, 1, 4).merge().setBackground('#d9d9d9'); vis.getRange(r, 3).setBackground('#d9d9d9'); }
         if (activeGrid[i].t === '9:30-9:45') { vis.getRange(r, 3, 1, 5).merge().setBackground('#d9d9d9'); }
         if (activeGrid[i].t === '11:45-1:00') { vis.getRange(r, 3, 1, 5).merge().setBackground('#d9d9d9'); }
      }

      currentRow += outputGrid.length + 2;
  });

  ss.toast('Visualizer generated successfully!', '✅ Done', 4);
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

    const data = termSheet.getRange(3, 1, Math.max(1, termSheet.getLastRow() - 2), 13).getValues();

    // Filter matching rows
    let rows = [];
    if (mode === 'Teacher') {
      rows = data.filter(r => r[2] && r[2].toString().toLowerCase() === targetName.toLowerCase());
    } else {
      rows = data.filter(r => r[0] && r[0].toString().toLowerCase() === targetName.toLowerCase());
    }

    if (!rows.length) return 'Error: No schedule found for ' + targetName;

    return buildPDFTemplateSheetAndExport(mode, targetName, rows);
  } catch (err) {
    return 'Error: ' + err.toString();
  }
}

function buildPDFTemplateSheetAndExport(mode, targetName, rows) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const tempName = 'TEMP_PDF_' + new Date().getTime();
  const ts = ss.insertSheet(tempName);

  try {
    if (mode === 'Teacher') {
      ts.getRange('A1:G1').merge().setValue('Name of Teacher: ' + targetName.toUpperCase()).setFontWeight('bold').setFontSize(14);
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

    const headers = ['TIME', 'NO. OF MINUTES', 'MONDAY', 'TUESDAY', 'WEDNESDAY', 'THURSDAY', 'FRIDAY'];
    ts.getRange('A4:G4').setValues([headers]).setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle').setBorder(true,true,true,true,true,true);
    ts.setColumnWidth(1, 100); ts.setColumnWidth(2, 90); ts.setColumnWidths(3, 5, 130);

    const grid = [
      { t: '7:00-7:30', m: 30, mon: 'Flag Raising\nCeremony', tue: 'GROUND PREPARATION/DAILY MORNING ROUTINE', wed: '', thu: '', fri: '', mergeSubj: true },
      { t: '7:30-8:30', m: 60, mon: '', tue: '', wed: '', thu: '', fri: '' },
      { t: '8:30-9:30', m: 60, mon: '', tue: '', wed: '', thu: '', fri: '' },
      { t: '9:30-9:45', m: 15, mon: 'RECESS/GROUP HANDWASHING', tue: '', wed: '', thu: '', fri: '', mergeSubj: true },
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

    let outputGrid = [];

    const getSchedule = (timeStr, dayIndex) => {
       const [inStr, outStr] = timeStr.split('-');
       const sStart = parseTime(inStr.includes('AM') || inStr.includes('PM') ? inStr : inStr + (parseInt(inStr.split(':')[0]) < 7 || parseInt(inStr.split(':')[0]) === 12 ? ' PM' : ' AM'));

       const matched = rows.find(r => {
          if (r[3 + dayIndex] !== true) return false;
          const rStart = parseTime(r[8]);
          return Math.abs(rStart - sStart) < 15;
       });

       if (!matched) return '';

       if (mode === 'Teacher') {
         return matched[1] + '\n' + matched[0]; // Subject + Grade Level
       } else {
         return matched[1] + '\n(' + matched[2] + ')'; // Subject + Teacher
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

    for (let i = 0; i < activeGrid.length; i++) {
       const r = 5 + i;
       if (activeGrid[i].t === '7:00-7:30') { ts.getRange(r, 4, 1, 4).merge().setBackground('#d9d9d9'); ts.getRange(r, 3).setBackground('#d9d9d9'); }
       if (activeGrid[i].t === '9:30-9:45') { ts.getRange(r, 3, 1, 5).merge().setBackground('#d9d9d9'); }
       if (activeGrid[i].t === '11:45-1:00') { ts.getRange(r, 3, 1, 5).merge().setBackground('#d9d9d9'); }
    }

    SpreadsheetApp.flush();
    const url = ss.getUrl().replace(/edit$/, '') + 'export?exportFormat=pdf&format=pdf' +
      '&size=A4&portrait=false&fitw=true&sheetnames=false&printtitle=false&pagenumbers=false' +
      '&gridlines=false&fzr=false&gid=' + ts.getSheetId();

    const token = ScriptApp.getOAuthToken();
    const response = UrlFetchApp.fetch(url, { headers: { 'Authorization': 'Bearer ' + token } });

    const blob = response.getBlob().setName(targetName + '_Schedule.pdf');
    const file = DriveApp.createFile(blob);

    ss.deleteSheet(ts);

    return file.getUrl();
  } catch (err) {
    if (ts) ss.deleteSheet(ts);
    return 'Error creating PDF: ' + err.toString();
  }
}
