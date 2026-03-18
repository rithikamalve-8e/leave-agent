import { buildAndSaveMonthlySummary, runYearStartAccrual } from "./excelManager";

// ── Helpers ────────────────────────────────────────────────────────────────

/** Returns true if date is a weekend */
function isWeekend(date: Date): boolean {
  return date.getDay() === 0 || date.getDay() === 6;
}

/** Returns the last working day of a given month */
function getLastWorkingDay(year: number, month: number): Date {
  // Start from last day of month and walk back until weekday
  // month is 1-based here, so we use month (not month-1) in Date constructor
  // because Date(year, month, 0) gives last day of previous month
  // i.e. Date(2026, 4, 0) = April 30 2026
  const lastDay = new Date(year, month, 0);
  while (isWeekend(lastDay)) {
    lastDay.setDate(lastDay.getDate() - 1);
  }
  return lastDay;
}

/** Returns ms until a specific time on a specific date */
function msUntil(target: Date): number {
  return Math.max(0, target.getTime() - Date.now());
}

/** Formats date as readable string for logs */
function formatDate(d: Date): string {
  return d.toLocaleString("en-IN", {
    day:    "2-digit",
    month:  "short",
    year:   "numeric",
    hour:   "2-digit",
    minute: "2-digit",
  });
}

// ── Midnight scheduler ─────────────────────────────────────────────────────
// Runs a callback every day at midnight + offset seconds
// Recursively reschedules itself

function scheduleDailyAt(
  hour:     number,
  minute:   number,
  second:   number,
  label:    string,
  callback: () => void
): void {
  function scheduleNext(): void {
    const now    = new Date();
    const target = new Date(
      now.getFullYear(),
      now.getMonth(),
      now.getDate(),
      hour,
      minute,
      second
    );

    // If target already passed today → schedule for tomorrow
    if (target <= now) {
      target.setDate(target.getDate() + 1);
    }

    const ms = msUntil(target);
    console.log(`[Scheduler] ${label} scheduled at ${formatDate(target)} (in ${Math.round(ms / 60000)} min)`);

    setTimeout(() => {
      console.log(`[Scheduler] ${label} triggered at ${formatDate(new Date())}`);
      try {
        callback();
      } catch (err) {
        console.warn(`[Scheduler] ${label} failed:`, err);
      }
      scheduleNext(); // reschedule for next day
    }, ms);
  }

  scheduleNext();
}

// ── Month-end summary job ──────────────────────────────────────────────────
// Runs at 11:59pm on the last working day of every month
// Builds MonthlySummary rows for ALL employees for that month

function scheduleMonthEndSummary(): void {
  function scheduleNext(): void {
    const now   = new Date();
    const year  = now.getFullYear();
    const month = now.getMonth() + 1; // 1-12

    // Find last working day of current month at 23:59:00
    const lastWorkingDay = getLastWorkingDay(year, month);
    const target         = new Date(
      lastWorkingDay.getFullYear(),
      lastWorkingDay.getMonth(),
      lastWorkingDay.getDate(),
      23, 59, 0
    );

    // If already passed this month → schedule for next month
    if (target <= now) {
      const nextMonth     = month === 12 ? 1    : month + 1;
      const nextYear      = month === 12 ? year + 1 : year;
      const nextLastDay   = getLastWorkingDay(nextYear, nextMonth);
      target.setFullYear(nextLastDay.getFullYear());
      target.setMonth(nextLastDay.getMonth());
      target.setDate(nextLastDay.getDate());
    }

    const ms = msUntil(target);
    console.log(`[Scheduler] Month-end summary scheduled at ${formatDate(target)} (in ${Math.round(ms / 60000)} min)`);

    setTimeout(() => {
      const runTime = new Date();
      console.log(`[Scheduler] Month-end summary triggered at ${formatDate(runTime)}`);

      try {
        const runMonth = runTime.getMonth() + 1; // 1-12
        const runYear  = runTime.getFullYear();

        console.log(`[Scheduler] Building summary for ${String(runMonth).padStart(2, "0")}/${runYear}`);
        buildAndSaveMonthlySummary(runMonth, runYear);
        console.log(`[Scheduler] Month-end summary complete`);
      } catch (err) {
        console.warn(`[Scheduler] Month-end summary failed:`, err);
      }

      scheduleNext(); // reschedule for next month
    }, ms);
  }

  scheduleNext();
}

// ── Jan 1st accrual job ────────────────────────────────────────────────────
// Runs at 00:00:05 on Jan 1st every year
// Sets carry_forward for all employees

function scheduleYearStartAccrual(): void {
  function scheduleNext(): void {
    const now         = new Date();
    const currentYear = now.getFullYear();

    // Next Jan 1st at 00:00:05
    let targetYear = currentYear;
    const jan1     = new Date(targetYear, 0, 1, 0, 0, 5); // Jan=0 in JS

    // If Jan 1st already passed this year → next year
    if (jan1 <= now) {
      targetYear++;
    }

    const target = new Date(targetYear, 0, 1, 0, 0, 5);
    const ms     = msUntil(target);

    console.log(`[Scheduler] Year-start accrual scheduled at ${formatDate(target)} (in ${Math.round(ms / 3600000)} hours)`);

    setTimeout(() => {
      console.log(`[Scheduler] Year-start accrual triggered at ${formatDate(new Date())}`);
      try {
        runYearStartAccrual();
        console.log(`[Scheduler] Year-start accrual complete`);
      } catch (err) {
        console.warn(`[Scheduler] Year-start accrual failed:`, err);
      }
      scheduleNext(); // reschedule for next Jan 1st
    }, ms);
  }

  scheduleNext();
}

// ── Month-end approver reminder ────────────────────────────────────────────
// Runs at 9:00am on the last working day of every month
// Sends DM to approvers listing their unactioned pending requests

function scheduleApproverReminder(
  sendReminderCallback: (month: number, year: number) => Promise<void>
): void {
  function scheduleNext(): void {
    const now   = new Date();
    const year  = now.getFullYear();
    const month = now.getMonth() + 1;

    const lastWorkingDay = getLastWorkingDay(year, month);
    const target         = new Date(
      lastWorkingDay.getFullYear(),
      lastWorkingDay.getMonth(),
      lastWorkingDay.getDate(),
      9, 0, 0   // 9am
    );

    if (target <= now) {
      const nextMonth   = month === 12 ? 1    : month + 1;
      const nextYear    = month === 12 ? year + 1 : year;
      const nextLastDay = getLastWorkingDay(nextYear, nextMonth);
      target.setFullYear(nextLastDay.getFullYear());
      target.setMonth(nextLastDay.getMonth());
      target.setDate(nextLastDay.getDate());
      target.setHours(9, 0, 0);
    }

    const ms = msUntil(target);
    console.log(`[Scheduler] Approver reminder scheduled at ${formatDate(target)} (in ${Math.round(ms / 60000)} min)`);

    setTimeout(async () => {
      const runTime = new Date();
      console.log(`[Scheduler] Approver reminder triggered at ${formatDate(runTime)}`);
      try {
        await sendReminderCallback(runTime.getMonth() + 1, runTime.getFullYear());
        console.log(`[Scheduler] Approver reminder complete`);
      } catch (err) {
        console.warn(`[Scheduler] Approver reminder failed:`, err);
      }
      scheduleNext();
    }, ms);
  }

  scheduleNext();
}

// ── HR takeover job ────────────────────────────────────────────────────────
// Runs at 6:00pm on the last working day of every month
// Any requests still Pending → escalate to HR

function scheduleHRTakeover(
  hrTakeoverCallback: (month: number, year: number) => Promise<void>
): void {
  function scheduleNext(): void {
    const now   = new Date();
    const year  = now.getFullYear();
    const month = now.getMonth() + 1;

    const lastWorkingDay = getLastWorkingDay(year, month);
    const target         = new Date(
      lastWorkingDay.getFullYear(),
      lastWorkingDay.getMonth(),
      lastWorkingDay.getDate(),
      18, 0, 0  // 6pm
    );

    if (target <= now) {
      const nextMonth   = month === 12 ? 1    : month + 1;
      const nextYear    = month === 12 ? year + 1 : year;
      const nextLastDay = getLastWorkingDay(nextYear, nextMonth);
      target.setFullYear(nextLastDay.getFullYear());
      target.setMonth(nextLastDay.getMonth());
      target.setDate(nextLastDay.getDate());
      target.setHours(18, 0, 0);
    }

    const ms = msUntil(target);
    console.log(`[Scheduler] HR takeover scheduled at ${formatDate(target)} (in ${Math.round(ms / 60000)} min)`);

    setTimeout(async () => {
      const runTime = new Date();
      console.log(`[Scheduler] HR takeover triggered at ${formatDate(runTime)}`);
      try {
        await hrTakeoverCallback(runTime.getMonth() + 1, runTime.getFullYear());
        console.log(`[Scheduler] HR takeover complete`);
      } catch (err) {
        console.warn(`[Scheduler] HR takeover failed:`, err);
      }
      scheduleNext();
    }, ms);
  }

  scheduleNext();
}

// ── Startup check ─────────────────────────────────────────────────────────
// Runs once on bot startup
// Checks if Jan 1st accrual was missed (bot was down on Jan 1st)

export function runStartupChecks(): void {
  const now         = new Date();
  const currentYear = now.getFullYear();

  console.log(`[Scheduler] Running startup checks...`);

  // Check if year-start accrual needs to run
  // runYearStartAccrual() internally checks year_entitlement_start
  // and skips employees already processed — safe to call always
  if (now.getMonth() === 0) {
    // It's January — make sure accrual ran
    console.log(`[Scheduler] January detected — checking year-start accrual`);
    runYearStartAccrual();
  }

  // Check if last month's summary was missed
  // e.g. bot was down on last working day of previous month
  const prevMonth = now.getMonth() === 0 ? 12 : now.getMonth(); // getMonth() is 0-based
  const prevYear  = now.getMonth() === 0 ? currentYear - 1 : currentYear;
  const prevLastWorkingDay = getLastWorkingDay(prevYear, prevMonth);

  // If today is AFTER the last working day of prev month
  // and it's within first 5 days of current month → likely missed
  if (now.getDate() <= 5) {
    console.log(
      `[Scheduler] Checking if ${String(prevMonth).padStart(2, "0")}/${prevYear}` +
      ` summary was missed...`
    );
    buildAndSaveMonthlySummary(prevMonth, prevYear);
  }

  console.log(`[Scheduler] Startup checks complete`);
}

// ── Main export — start all schedulers ────────────────────────────────────

export function startSchedulers(callbacks: {
  sendApproverReminder: (month: number, year: number) => Promise<void>;
  sendHRTakeover:       (month: number, year: number) => Promise<void>;
}): void {

  console.log(`[Scheduler] Starting all schedulers...`);

  // 1. Month-end summary at 11:59pm last working day
  scheduleMonthEndSummary();

  // 2. Jan 1st carry-forward accrual
  scheduleYearStartAccrual();

  // 3. Approver reminder at 9am last working day
  scheduleApproverReminder(callbacks.sendApproverReminder);

  // 4. HR takeover at 6pm last working day
  scheduleHRTakeover(callbacks.sendHRTakeover);

  console.log(`[Scheduler] All schedulers started`);
}

// ── Manual trigger for testing ─────────────────────────────────────────────
// Call this from devtools: "build summary march 2026"

export function triggerMonthlySummaryNow(monthNum: number, year: number): void {
  console.log(`[Scheduler] Manual trigger: building summary for ${String(monthNum).padStart(2, "0")}/${year}`);
  buildAndSaveMonthlySummary(monthNum, year);
}