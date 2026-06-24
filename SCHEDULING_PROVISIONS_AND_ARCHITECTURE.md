# Autonomous Faculty Loading System — Provisions & Architecture

This document serves as the strict rulebook and reference map for the Google Apps Script project governing the Autonomous Faculty Loading System. It must be consulted for any future revisions, feature additions, or debugging.

## The Architecture Map (Local CSP Solver)

The system utilizes a 100% local JavaScript execution environment, relying on an advanced Constraint Satisfaction Problem (CSP) solver with chronological backtracking.

1. **Data Extraction:** The \`runAutoScheduler\` function acts as the central engine. It reads master arrays from:
   - \`SECTION_ENROLL\`
   - \`SUBJECT_LOAD\` (which includes the assigned teacher and weekly hours)
   - \`TEACHER_ENROLL\`
2. **Variable Ordering & Load Balancing:**
   - Unassigned subjects (\`⚠️ Unassigned\`) are dynamically mapped to available teachers sharing the same specialization who possess the lowest current load.
   - High-priority constraints (Homerooms, ARAL) are pre-mapped to the matrix before loop execution.
3. **Chronological Backtracking Engine:**
   - The script iterates through a strict timeline. Before a block is committed to the matrix, it undergoes Look-Ahead Validation to ensure its duration (30m, 60m, 90m) does not bleed across physical breaks.
   - If a constraint is mathematically violated (e.g., Teacher Daily Hard Limit), the engine penalizes the score and recursively attempts the next permutation, advancing the chronological tracker (\`time += 30\`) to force dense packing.
4. **Data Injection:** The engine maps the solved true/false states into the 13-column Term tab format and utilizes a rapid \`setValues\` injection to render the final schedule.

## The Constraints Ledger

The Local CSP Engine evaluates the following hard and soft constraints natively inside the \`runAutoScheduler\` logic block:

- **Lunch Break:** Strictly locked between **11:45 AM - 1:00 PM**. No classes may be scheduled during this block.
- **Recess Break:** Strictly locked between **9:30 AM - 9:45 AM**. No classes may be scheduled during this block.
- **ARAL Subject:** Must be strictly plotted exactly at **3:00 PM - 4:00 PM**.
- **Block Duration Logic:**
  - **Homeroom:** Exactly **30 minutes**.
  - **Phil Gov (Philippine Politics):** Exactly **90 minutes**.
  - **Standard Subjects:** Minimum of **60 minutes**. If fractional weekly hours exist (e.g., 2.5 hours), the remainder must utilize a **90-minute** block to bridge the fraction.
- **Teacher Daily Limits:**
  - Hard Limit: Maximum of **6 hours** of teaching per day.
  - Soft Limit: Preferred maximum of **4.5 hours** per day. (If exceeded, the warning \`🟠 >4.5h Soft Limit\` must be flagged).
- **Junior High School (JHS / G7-10) Rules:**
  - Sections **cannot** meet the same subject twice in one day.
  - The Homeroom Adviser **cannot** be booked/conflict during the section's other scheduled classes.
- **Senior High School (SHS / G11-12) Rules:**
  - Sections **can** meet the same subject twice in one day.
  - However, the identical subjects **cannot be consecutive** (no back-to-back blocks).
  - Homeroom is strictly plotted at **3:00 PM**.

## ⚠️ Isolation Protocol

**The dashboards (Teacher and Section), the All Sections Visualizer, the manual Conflict Checkers (\`onEdit\`), and the PDF Generation scripts are container-bound UI elements.**

**DO NOT ALTER THESE FUNCTIONS WHEN TWEAKING THE \`runAutoScheduler\` CSP ENGINE.**

The UI elements depend on the final 13-column output structure of the Term tabs. As long as the solver accurately populates those columns via \`setValues\`, the visualizers and checkers will function autonomously.
