# Autonomous Faculty Loading System — Provisions & Architecture

This document serves as the strict rulebook and reference map for the Google Apps Script project governing the Autonomous Faculty Loading System. It must be consulted for any future revisions, feature additions, or debugging.

## The Architecture Map (Dynamic Background Guide)

The system has shifted from a bulk "Auto-Scheduler" approach to a highly interactive, dynamic, and background-synchronized manual encoding environment.

1. **Manual Encoding with Dynamic Allowances:** The data validations across all tabs have been modified to \`setAllowInvalid(true)\`. This means the system will strictly validate and warn about rules but will **never physically block** a human encoder from writing data.
2. **Real-Time Provision Feedback (\`checkRowConflicts\`):** This is the core dynamic guide. Every time a row is edited in a Term tab, this function evaluates the exact limits (Lunch, Recess, durations, ARAL/Homeroom placement, 6-hour max) and instantly outputs actionable error/warning strings into the "Warnings" column.
3. **Background Dashboard Synchronization:**
   - Upon any edit to the Term tabs, the system automatically triggers the Master Conflict Checker (\`runConflictChecker\`) in the background.
   - It also automatically rebuilds the \`buildAllSectionsVisualizer\` and the new \`buildAllTeachersVisualizer\` tabs silently in the background, ensuring all graphical views are instantly updated as the operator encodes.

## The Constraints Ledger (Warnings System)

The Dynamic Guide natively evaluates and outputs warnings for the following provisions:

- **Lunch Break:** Strictly locked between **11:45 AM - 1:00 PM**. Classes bleeding into this generate a \`🔴 Bleeds into Lunch\` error.
- **Recess Break:** Strictly locked between **9:30 AM - 9:45 AM**. Classes bleeding into this generate a \`🔴 Bleeds into Recess\` error.
- **ARAL Subject:** Must be strictly plotted exactly at **3:00 PM**.
- **Block Duration Logic:**
  - **Homeroom:** Exactly **30 minutes**.
  - **Phil Gov (Philippine Politics):** Exactly **90 minutes**.
  - **Standard Subjects:** Minimum of **60 minutes**.
- **Teacher Daily Limits:**
  - Hard Limit: Maximum of **6 hours** of teaching per day. Generates a \`🔴 Teacher >6h daily limit\` error.
- **Junior High School (JHS / G7-10) Rules:**
  - Sections **cannot** meet the same subject twice in one day. (\`🔴 JHS: Met twice in a day\`)
  - Homeroom must be on Monday at **7:30 AM**. (\`🟠 JHS Homeroom should be Mon 7:30 AM\`)
- **Senior High School (SHS / G11-12) Rules:**
  - Identical subjects **cannot be consecutive** (no back-to-back blocks). (\`🔴 SHS: Back-to-back same subject\`)
  - Homeroom is strictly plotted at **3:00 PM**. (\`🟠 SHS Homeroom should be 3:00 PM\`)

## ⚠️ Isolation Protocol

**The dashboards (Teacher and Section), the All Sections/Teachers Visualizers, and the PDF Generation scripts are container-bound UI elements.**

These elements are designed to run autonomously and silently in the background. DO NOT break their structural loop or the \`onEdit\` trigger mechanisms.
