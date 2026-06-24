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
    .addItem('6️⃣ Build: AI Feedback Workspace', 'buildAIFeedbackWorkspace')
    .addItem('7️⃣ Build: All Sections Visualizer', 'buildAllSectionsVisualizer')
    .addSeparator()
    .addItem('🚀 RUN OPENROUTER AUTO-SCHEDULER', 'runAutoScheduler')
    .addItem('🤖 VALIDATE AI PROVISION', 'validateProvisionWithOpenRouter')
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

  ss.toast('Packaging data for OpenRouter AI...', '🧠 Computing', 4);

  const getSafeData = (sheet, numCols) => {
    const lr = sheet.getLastRow();
    return lr < 3 ? [] : sheet.getRange(3, 1, lr - 2, numCols).getValues();
  };

  const demands = getSafeData(subSheet, 5).filter(r => r[0] && r[1] && r[2]); // Col E is Assigned Teacher
  const teachers = getSafeData(enrollSheet, 3).filter(r => r[1]);

  CFG.TERMS.forEach(term => {
    const termDemands = demands.filter(d => d[0] === term);
    if (!termDemands.length) return;

    // Clear existing tab data
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

    const payload = {
      term: term,
      demands: termDemands.map(d => ({
        section: d[1].toString().trim(),
        subject: d[2].toString().trim(),
        weeklyHours: parseFloat(d[3]) || 0,
        assignedTeacher: d[4] ? d[4].toString().trim() : '⚠️ Unassigned'
      })),
      teachers: teachers.map(t => ({
        name: t[1].toString().trim(),
        specialization: t[2] ? t[2].toString().trim() : 'None'
      }))
    };

    fetchOpenRouterSchedule(ss, term, payload);
  });

  if (typeof updateSubjectLoadingHours === 'function') updateSubjectLoadingHours();
}

function fetchOpenRouterSchedule(ss, term, payload) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('OPENROUTER_API_KEY');
  if (!apiKey) {
    return ss.toast('Error: OPENROUTER_API_KEY missing in Script Properties. Add your key.', '⚠️ API Error', 10);
  }

  const systemPrompt = `You are an expert school auto-scheduler. I am providing you with the teaching demands and available teachers.

You must generate a schedule that perfectly maps every subject requirement to the following 13 columns exactly as expected in the Google Sheet:
["Section", "Subject", "Teacher", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Time In", "Time Out", "Action", "Warnings", "Suggested Teachers"]

CRITICAL SCHEDULING CONSTRAINTS:
1. Lunch: Strictly 11:45 AM - 1:00 PM (No classes allowed).
2. Recess: Strictly 9:30 AM - 9:45 AM (No classes allowed).
3. ARAL Subject: Must be strictly placed at 3:00 PM - 4:00 PM.
4. Homeroom Duration: Exactly 30 minutes.
5. Phil Gov Subject Duration: Exactly 90 minutes.
6. Standard Subjects Duration: Minimum 60 minutes. Fractional weekly hours (e.g., 2.5) should utilize a 90-minute block for the remainder.
7. Teacher Daily Limits: Maximum 6 hours daily hard limit. Soft limit of 4.5 hours (flag with "🟠 >4.5h Soft Limit" in Warnings if exceeded).
8. JHS Constraints (Grades 7-10): Sections cannot meet the same subject twice in one day. The Homeroom Adviser cannot be booked during the section's other classes.
9. SHS Constraints (Grades 11-12): Sections can meet the same subject twice in one day, but NOT consecutively (no back-to-back blocks). Homeroom is strictly at 3:00 PM (Resolve any ARAL vs Homeroom 3:00 PM collisions logically, e.g., ARAL at 3:00 PM on other days, Homeroom on Monday).
10. Unassigned Subjects: Distribute them to teachers with low loads based on specialization, flag with "🟢 Distributed Sub" in Warnings.
11. If a teacher is overloaded, fallback to "Unavailable Teacher" and suggest free teachers in "Suggested Teachers".
12. For Monday, Tuesday, Wednesday, Thursday, Friday, output boolean true or false.
13. Time In and Time Out should be formatted as "h:mm AM/PM" (e.g., "7:30 AM").
14. Action is always "—".

Return ONLY a raw JSON object with a single root key "schedule" containing an array of objects. Each object must have keys matching the 13 columns exactly (case-sensitive).`;

  const userPrompt = `Payload for ${term}: ${JSON.stringify(payload)}`;

  const apiPayload = {
    "model": "openrouter/free",
    "response_format": { "type": "json_object" },
    "messages": [
      { "role": "system", "content": systemPrompt },
      { "role": "user", "content": userPrompt }
    ]
  };

  const options = {
    'method': 'post',
    'contentType': 'application/json',
    'headers': {
      'Authorization': 'Bearer ' + apiKey
    },
    'payload': JSON.stringify(apiPayload),
    'muteHttpExceptions': true
  };

  try {
    const url = 'https://openrouter.ai/api/v1/chat/completions';
    const response = UrlFetchApp.fetch(url, options);
    const responseText = response.getContentText();

    // Check for HTTP errors before parsing
    if (response.getResponseCode() !== 200) {
        throw new Error('HTTP ' + response.getResponseCode() + ': ' + responseText);
    }

    const json = JSON.parse(responseText);

    if (json.choices && json.choices.length > 0) {
      let aiResponseText = json.choices[0].message.content;

      // Attempt to clean markdown block formatting if the AI includes it despite instructions
      if (aiResponseText.startsWith('```json')) {
          aiResponseText = aiResponseText.replace(/^```json\n/, '').replace(/\n```$/, '');
      } else if (aiResponseText.startsWith('```')) {
          aiResponseText = aiResponseText.replace(/^```\n/, '').replace(/\n```$/, '');
      }

      let scheduleData = JSON.parse(aiResponseText).schedule;

      if (scheduleData && scheduleData.length > 0) {
          writeScheduleToSheet(ss, term, scheduleData);
          ss.toast(`Successfully generated ${term} schedule using OpenRouter.`, '✅ Success', 5);
      } else {
          ss.toast(`OpenRouter returned empty schedule for ${term}.`, '⚠️ Warning', 5);
      }
    } else {
      ss.toast('Error parsing OpenRouter response: ' + responseText, '⚠️ API Error', 10);
    }
  } catch (e) {
    ss.toast('API Request Failed: ' + e.toString(), '⚠️ Error', 10);
  }
}

function writeScheduleToSheet(ss, term, scheduleData) {
  const ts = ss.getSheetByName(term);
  if (!ts) return;

  const outputRows = scheduleData.map(row => [
    row["Section"], row["Subject"], row["Teacher"],
    row["Monday"], row["Tuesday"], row["Wednesday"], row["Thursday"], row["Friday"],
    row["Time In"], row["Time Out"], row["Action"], row["Warnings"], row["Suggested Teachers"]
  ]);

  if (outputRows.length > 0) {
      ts.getRange(3, 1, outputRows.length, 13).setValues(outputRows);
      ts.getRange(3, 1, outputRows.length, 13).setVerticalAlignment('middle');
      ts.getRange(3, 9, outputRows.length, 2).setHorizontalAlignment('center');
      ts.getRange(3, 12, outputRows.length, 1).setFontColor(C.warn).setFontStyle('italic');
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

  if (!secSheet) return;
  const rule = SpreadsheetApp.newDataValidation().requireValueInRange(secSheet.getRange('B3:B1000'), true).build();

  if (subSheet) subSheet.getRange('B3:B1000').setDataValidation(rule);

  CFG.TERMS.forEach(term => {
    const termSheet = ss.getSheetByName(term);
    if (termSheet) termSheet.getRange('A3:A1000').setDataValidation(rule);
  });
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

  CFG.TERMS.forEach(term => {
    const termSheet = ss.getSheetByName(term);
    if (termSheet) termSheet.getRange('C3:C1000').setDataValidation(rule);
  });
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
    sheet.getRange('A1:M1').merge().setValue(`📅 ${term.toUpperCase()} MASTER SCHEDULE`).setFontWeight('bold').setFontSize(12).setBackground(C.navyDark).setFontColor(C.white).setHorizontalAlignment('center').setVerticalAlignment('middle');
    sheet.setRowHeight(1, 40);

    const headers = ['Section', 'Subject', 'Teacher', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Time In', 'Time Out', 'Action', 'Warnings', 'Suggested Teachers'];
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

  updateSectionDropdowns();
  updateTeacherNameDropdowns();
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


function buildAIFeedbackWorkspace() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const wsName = 'AI Feedback Workspace';
  let ws = ss.getSheetByName(wsName) || ss.insertSheet(wsName, 6);
  ws.clear();

  ws.getRange('A1:B1').setValues([['Proposed Scheduling Rule', 'AI Evaluation']]).setFontWeight('bold').setFontSize(12).setBackground(C.navyDark).setFontColor(C.white).setHorizontalAlignment('center').setVerticalAlignment('middle');
  ws.setRowHeight(1, 40);

  ws.setColumnWidth(1, 500);
  ws.setColumnWidth(2, 500);

  ws.getRange('A2:B2').setWrap(true).setVerticalAlignment('top');
  ws.setRowHeight(2, 200);

  ws.getRange('A2').setBackground(C.tealLight).setBorder(true, true, true, true, null, null);
  ws.getRange('B2').setBackground(C.navyLight).setBorder(true, true, true, true, null, null);

  ss.toast('AI Feedback Workspace created.', '✅ Phase 6 Complete', 4);
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

function validateProvisionWithOpenRouter() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName('AI Feedback Workspace');
  if (!ws) {
    return ss.toast('AI Feedback Workspace not found. Build it from the Faculty Tools menu.', '⚠️ Error');
  }

  const rule = ws.getRange('A2').getValue();
  if (!rule) {
    return ss.toast('Please enter a proposed scheduling rule in A2.', '⚠️ Empty');
  }

  ws.getRange('B2').setValue('⏳ Thinking...').setFontColor(C.muted);

  const apiKey = PropertiesService.getScriptProperties().getProperty('OPENROUTER_API_KEY');
  if (!apiKey) {
    ws.getRange('B2').setValue('Error: OPENROUTER_API_KEY missing in Script Properties.');
    return;
  }

  const systemPrompt = "You are an expert school scheduling analyst. Evaluate the rule provided by the user against standard school scheduling logic. You must reply strictly with either the word EXECUTABLE or IMPOSSIBLE, followed by a colon and exactly two sentences explaining potential bottlenecks or implications.";
  const userPrompt = "Analyze the following proposed scheduling rule:\n\n\"" + rule + "\"";

  const payload = {
    "model": "openrouter/free",
    "messages": [
      { "role": "system", "content": systemPrompt },
      { "role": "user", "content": userPrompt }
    ]
  };

  const options = {
    'method': 'post',
    'contentType': 'application/json',
    'headers': {
      'Authorization': 'Bearer ' + apiKey
    },
    'payload': JSON.stringify(payload),
    'muteHttpExceptions': true
  };

  try {
    const url = 'https://openrouter.ai/api/v1/chat/completions';
    const response = UrlFetchApp.fetch(url, options);

    if (response.getResponseCode() !== 200) {
        ws.getRange('B2').setValue('HTTP Error: ' + response.getContentText());
        return;
    }

    const json = JSON.parse(response.getContentText());

    if (json.choices && json.choices.length > 0) {
      let aiText = json.choices[0].message.content;
      ws.getRange('B2').setValue(aiText.trim()).setFontColor(C.body);
    } else {
      ws.getRange('B2').setValue('Error parsing AI response: ' + response.getContentText());
    }
  } catch (e) {
    ws.getRange('B2').setValue('API Error: ' + e.toString());
  }
}
