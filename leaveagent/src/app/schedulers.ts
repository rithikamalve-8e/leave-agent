/**
 * schedulers.ts
 * All time-based jobs for LeaveAgent.
 * - Daily 9am workforce summary → announcement channel
 * - Month-end approver reminder (last working day 9am)
 * - Month-end HR takeover (last working day 6pm)
 * - Dec 25 carry-forward calculation
 */

import { getTodaysAbsences, getAllLeaveRequests, getAllEmployees } from "./postgresManager";
import { sendDailySummaryRest } from "./notificationService";

// ── Helpers ────────────────────────────────────────────────────────────────

function isWeekend(date: Date): boolean {
  return date.getDay() === 0 || date.getDay() === 6;
}

function lastWorkingDayOfMonth(year: number, month: number): Date {
  const last = new Date(year, month, 0); // last day of month
  while (isWeekend(last)) last.setDate(last.getDate() - 1);
  return last;
}

function msUntil(target: Date): number {
  return Math.max(0, target.getTime() - Date.now());
}

function formatDate(d: Date): string {
  return d.toLocaleString("en-IN", { dateStyle: "medium", timeStyle: "short" });
}

// Safe setTimeout that handles delays > 2^31 ms (Node limit)
function safeSetTimeout(fn: () => void, ms: number): void {
  const MAX = 2 ** 31 - 1;
  if (ms <= MAX) { setTimeout(fn, ms); return; }
  setTimeout(() => safeSetTimeout(fn, ms - MAX), MAX);
}

// ── Daily 9am Summary ──────────────────────────────────────────────────────

function scheduleDailySummary(): void {
  const announcementConvId = process.env.ANNOUNCEMENT_CHANNEL_ID;
  if (!announcementConvId) {
    console.log("[Scheduler] No ANNOUNCEMENT_CHANNEL_ID — daily summary disabled.");
    return;
  }

  async function postDailySummary(): Promise<void> {
    const today = new Date();
    if (isWeekend(today)) return;

    const records = await getTodaysAbsences();
    if (records.length === 0) return;

    await sendDailySummaryRest(records as any[]);
  }

  function scheduleNext(): void {
    const now     = new Date();
    const next9am = new Date(now);
    next9am.setHours(9, 0, 0, 0);
    if (now >= next9am) next9am.setDate(next9am.getDate() + 1);

    const ms = msUntil(next9am);
    const h  = Math.floor(ms / 3600000);
    const m  = Math.floor((ms % 3600000) / 60000);
    console.log(`[Scheduler] Daily summary scheduled — next run in ${h}h ${m}m (${formatDate(next9am)})`);

    safeSetTimeout(async () => {
      try { await postDailySummary(); } catch (err) { console.warn("[Scheduler] Daily summary failed:", err); }
      scheduleNext();
    }, ms);
  }

  scheduleNext();
}

// ── Month-End Approver Reminder (last working day 9am) ────────────────────

function scheduleApproverReminder(
  callback: (month: number, year: number) => Promise<void>
): void {
  function scheduleNext(): void {
    const now   = new Date();
    const year  = now.getFullYear();
    const month = now.getMonth();

    // Try current month's last working day
    let target = lastWorkingDayOfMonth(year, month + 1);
    target.setHours(9, 0, 0, 0);

    // If already past, schedule for next month
    if (now >= target) {
      const nextMonth = month + 1 > 11 ? 0 : month + 1;
      const nextYear  = month + 1 > 11 ? year + 1 : year;
      target = lastWorkingDayOfMonth(nextYear, nextMonth + 1);
      target.setHours(9, 0, 0, 0);
    }

    const ms = msUntil(target);
    console.log(`[Scheduler] Approver reminder scheduled for ${formatDate(target)} (in ${Math.round(ms / 60000)} min)`);

    safeSetTimeout(async () => {
      const runTime = new Date();
      try {
        await callback(runTime.getMonth() + 1, runTime.getFullYear());
      } catch (err) {
        console.warn("[Scheduler] Approver reminder failed:", err);
      }
      scheduleNext();
    }, ms);
  }

  scheduleNext();
}

// ── Month-End HR Takeover (last working day 6pm) ──────────────────────────

function scheduleHRTakeover(
  callback: (month: number, year: number) => Promise<void>
): void {
  function scheduleNext(): void {
    const now   = new Date();
    const year  = now.getFullYear();
    const month = now.getMonth();

    let target = lastWorkingDayOfMonth(year, month + 1);
    target.setHours(18, 0, 0, 0);

    if (now >= target) {
      const nextMonth = month + 1 > 11 ? 0 : month + 1;
      const nextYear  = month + 1 > 11 ? year + 1 : year;
      target = lastWorkingDayOfMonth(nextYear, nextMonth + 1);
      target.setHours(18, 0, 0, 0);
    }

    const ms = msUntil(target);
    console.log(`[Scheduler] HR takeover scheduled for ${formatDate(target)} (in ${Math.round(ms / 60000)} min)`);

    safeSetTimeout(async () => {
      const runTime = new Date();
      try {
        await callback(runTime.getMonth() + 1, runTime.getFullYear());
      } catch (err) {
        console.warn("[Scheduler] HR takeover failed:", err);
      }
      scheduleNext();
    }, ms);
  }

  scheduleNext();
}

// ── Dec 25 Carry Forward ───────────────────────────────────────────────────

function scheduleDec25CarryForward(): void {
  function scheduleNext(): void {
    const now    = new Date();
    const year   = now.getFullYear();
    let   target = new Date(year, 11, 25, 9, 0, 0, 0); // Dec 25 9am

    if (now >= target) target = new Date(year + 1, 11, 25, 9, 0, 0, 0);

    const ms = msUntil(target);
    console.log(`[Scheduler] Dec 25 carry-forward scheduled for ${formatDate(target)}`);

    safeSetTimeout(async () => {
      try {
        await runCarryForward();
        console.log("[Scheduler] Dec 25 carry-forward complete");
      } catch (err) {
        console.warn("[Scheduler] Dec 25 carry-forward failed:", err);
      }
      scheduleNext();
    }, ms);
  }

  scheduleNext();
}

async function runCarryForward(): Promise<void> {
  const { PrismaClient } = await import("@prisma/client");
  const prisma  = new PrismaClient();
  const allEmps = await getAllEmployees();
  const year    = new Date().getFullYear();

  for (const emp of allEmps) {
    // Max 6 days carry forward
    const carry = Math.min(emp.leave_balance, 6);
    await prisma.employee.update({
      where: { name: emp.name },
      data:  { carry_forward: carry },
    });
    console.log(`[Scheduler] Carry forward: ${emp.name} → ${carry} days`);
  }

  await prisma.$disconnect();
}

// ── Monthly Summary snapshot ───────────────────────────────────────────────

export function buildAndSaveMonthlySummary(month: number, year: number): void {
  console.log(`[Scheduler] Monthly summary snapshot for ${String(month).padStart(2,"0")}/${year} — stored in Postgres MonthlySummary table`);
  // Actual snapshot written to MonthlySummary table by postgresManager when report is generated
}

// ── Startup Checks ─────────────────────────────────────────────────────────

export function runStartupChecks(): void {
  const now = new Date();
  console.log(`[Scheduler] Startup checks at ${formatDate(now)}`);

  // Check if we're on last working day and past 9am — reminder might have been missed
  const lastDay = lastWorkingDayOfMonth(now.getFullYear(), now.getMonth() + 1);
  const isLastWorkingDay = now.toDateString() === lastDay.toDateString();
  if (isLastWorkingDay && now.getHours() >= 9) {
    console.log("[Scheduler] Running on last working day — approver reminders may need manual trigger");
  }

  console.log("[Scheduler] Startup checks complete");
}

// ── Trigger manual summary ─────────────────────────────────────────────────

export function triggerMonthlySummaryNow(month: number, year: number): void {
  console.log(`[Scheduler] Manual trigger: ${String(month).padStart(2,"0")}/${year}`);
  buildAndSaveMonthlySummary(month, year);
}

// ── Main export ────────────────────────────────────────────────────────────

export function startSchedulers(callbacks: {
  sendApproverReminder: (month: number, year: number) => Promise<void>;
  sendHRTakeover:       (month: number, year: number) => Promise<void>;
}): void {
  console.log("[Scheduler] Starting all schedulers...");
  scheduleDailySummary();
  scheduleDec25CarryForward();
  scheduleApproverReminder(callbacks.sendApproverReminder);
  scheduleHRTakeover(callbacks.sendHRTakeover);
  console.log("[Scheduler] All schedulers started");
}
