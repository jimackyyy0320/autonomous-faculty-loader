# 🏫 Autonomous Faculty Loading System — Architecture & Rules

## System Overview
This is a Google Apps Script (GAS) project attached to a Google Sheet acting as a strict relational Management Information System (MIS). It autonomously maps out faculty schedules, detects constraints, and provides humanized dashboards to advance a supportive environment for our teachers and cohorts.

## Core Directives for Jules AI:
1. **Protect the Triggers:** The `onEdit(e)` trigger is the lifeblood of this system. Do NOT remove or decouple the dynamic dropdown syncs or cascading deletes.
2. **Maintain Single-File Architecture:** Keep all configurations, triggers, algorithms, and builders within this single `Code.js` file to ensure easy deployment.
3. **The "Humanized" Constraint:** Scheduling logic must prioritize empathy and workload balance. For example, JHS Homerooms are strictly locked to Monday at 7:30 AM to support a strong start to the week. Do not "optimize" these constraints away.
4. **Error Handling:** When hitting a constraint bottleneck, do not fail silently. Always utilize `SpreadsheetApp.getUi().alert()` to provide clear, diagnostic feedback to the user.
