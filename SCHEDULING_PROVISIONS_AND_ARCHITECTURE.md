# Autonomous Faculty Loading System — Provisions & Architecture

This document serves as the strict rulebook and reference map for the Google Apps Script project governing the Autonomous Faculty Loading System. It must be consulted for any future revisions, feature additions, or debugging.

## The Architecture Map

The system has transitioned from a deterministic dense-packing loop to an AI-driven inference engine utilizing the Gemini API. The flow of data is as follows:

1. **Data Extraction:** The \`runAutoScheduler\` function acts as the central data packager. It reads master arrays from:
   - \`SECTION_ENROLL\`
   - \`SUBJECT_LOAD\` (which includes the assigned teacher and weekly hours)
   - \`TEACHER_ENROLL\`
2. **Payload Construction:** The data is transformed into a lightweight JSON payload detailing the demands per term and the roster of available teachers (with their specializations).
3. **API Invocation:** The \`fetchGeminiSchedule\` function takes this payload and constructs a strict prompt containing all non-negotiable scheduling constraints.
   - It utilizes \`UrlFetchApp\` to make a POST request to the Gemini API.
   - **Crucial:** It enforces Structured Outputs using JSON mode (\`responseSchema\`) to guarantee the AI returns an array of objects perfectly matching the 13-column Term tab format.
   - The API key is securely retrieved using \`PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY')\`. **Never hardcode the key.**
4. **Data Injection:** The returned JSON is parsed. The \`writeScheduleToSheet\` function maps the objects back into 2D arrays and uses \`setValues\` to render the final schedule onto the respective Term tabs.

## The Constraints Ledger

The Gemini API prompt is explicitly fed the following hard and soft constraints. Any modification to how schedules are plotted must begin by updating these prompt rules:

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
  - Homeroom is strictly plotted at **3:00 PM**. (The AI must logically resolve ARAL vs. Homeroom collisions at this time slot).

## ⚠️ Isolation Protocol

**The dashboards (Teacher and Section), the All Sections Visualizer, the manual Conflict Checkers (\`onEdit\`), and the PDF Generation scripts are container-bound UI elements.**

**DO NOT ALTER THESE FUNCTIONS WHEN TWEAKING THE \`runAutoScheduler\` OR THE GEMINI INFERENCE ENGINE.**

The UI elements depend on the final 13-column output structure of the Term tabs. As long as the AI accurately populates those columns via \`setValues\`, the visualizers and checkers will function autonomously.
