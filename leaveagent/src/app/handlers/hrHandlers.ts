import * as XLSX from "xlsx";
import {
  buildAllRequestsCard,
  buildAuditLogCard,
  buildEmployeeProfileCard,
  buildLeaveBalanceCard,
  buildUnregisteredCard,
  buildOrgChartCard,
  buildUnactionedCard,
  buildSuccessCard,
  buildErrorCard,
  buildMyRequestsCard,
  buildHolidaysCard,
  formatDisplayDate,
} from "../cards";
import {
  getAllLeaveRequests,
  getLeaveRequestsByEmployee,
  getAbsencesForDateRange,
  getTodaysAbsences,
  findEmployee,
  getAllEmployees,
  updateLeaveStatus,
  deleteLeaveRequest,
  addLeaveRequest,
  adjustLeaveBalance,
  addHoliday,
  getHolidays,
  isHoliday,
  getAuditLog,
  getUnregisteredEmployees,
  getMonthlyPendingRequests,
  getLeaveBalance,
  appendAuditLog,
  upsertEmployee,
  clearAllHolidays,
} from "../postgresManager";
import {
  sendStatusCardToEmployee,
  sendApprovalAnnouncement,
  sendHRAlert,
  sendBalanceAdjustedNotification,
  sendDeleteNotifications,
  sendHolidayNotificationToAll,
  NotificationContext,
} from "../notificationServices";
import { CommandContext } from "./sharedHandlers";
import { handleWhoIsOnLeave } from "./approveHandlers";
import * as path from "path";
import * as fs   from "fs";

// ── Shared month/date helpers ──────────────────────────────────────────────

const MONTH_NAMES = [
  "january","february","march","april","may","june",
  "july","august","september","october","november","december",
];

/**
 * Returns { mm: "01"–"12", name: "January" } if a month name is found in
 * the lowercased command string, otherwise null.
 */
function extractMonth(cmd: string): { mm: string; name: string } | null {
  const lower = cmd.toLowerCase();
  for (let i = 0; i < MONTH_NAMES.length; i++) {
    if (lower.includes(MONTH_NAMES[i])) {
      return {
        mm:   String(i + 1).padStart(2, "0"),
        name: MONTH_NAMES[i].charAt(0).toUpperCase() + MONTH_NAMES[i].slice(1),
      };
    }
  }
  return null;
}

/**
 * Parses a YYYY-MM-DD date from the message, or returns today's date.
 */
function extractDate(cmd: string): string {
  const m = cmd.match(/(\d{4}-\d{2}-\d{2})/);
  return m ? m[1] : new Date().toISOString().split("T")[0];
}

// ── all requests ───────────────────────────────────────────────────────────

export async function handleAllRequests(ctx: CommandContext): Promise<void> {
  let records = await getAllLeaveRequests();

  // Filter by month name if present
  const monthInfo = extractMonth(ctx.cmd);
  if (monthInfo) {
    const yy = String(new Date().getFullYear());
    records   = records.filter((r) => r.date.startsWith(`${yy}-${monthInfo.mm}`));
  }

  // Filter by employee name: "all requests Rithika MR"
  // Strip leading command words to isolate the name portion
  const nameMatch = ctx.userMessage.match(/all\s+requests\s+(.+)/i);
  if (nameMatch) {
    const rest    = nameMatch[1].trim();
    const isMonth = extractMonth(rest) !== null;
    if (rest && !isMonth) {
      records = records.filter((r) =>
        r.employee.toLowerCase().includes(rest.toLowerCase())
      );
    }
  }

  const label = monthInfo
    ? `All Requests — ${monthInfo.name}`
    : "All Requests";

  await ctx.send(buildAllRequestsCard(records as any[], label));
}

// ── pending ────────────────────────────────────────────────────────────────

export async function handleAllPending(ctx: CommandContext): Promise<void> {
  let records = await getAllLeaveRequests();
  records = records.filter((r) => r.status === "Pending");

  // Optionally filter by employee name: "pending Rithika"
  const nameMatch = ctx.userMessage.match(/pending\s+(.+)/i);
  if (nameMatch) {
    const name = nameMatch[1].trim();
    records = records.filter((r) =>
      r.employee.toLowerCase().includes(name.toLowerCase())
    );
  }

  const nameStr = nameMatch ? ` — ${nameMatch[1].trim()}` : "";
  await ctx.send(buildAllRequestsCard(records as any[], `Pending Requests${nameStr}`));
}

// ── unactioned ─────────────────────────────────────────────────────────────

export async function handleUnactioned(ctx: CommandContext): Promise<void> {
  const records = await getMonthlyPendingRequests();
  await ctx.send(buildUnactionedCard(records as any[], true));
}

// ── approve all unactioned ─────────────────────────────────────────────────

export async function handleApproveAllUnactioned(
  ctx:  CommandContext,
  nctx: NotificationContext
): Promise<void> {
  const records = await getMonthlyPendingRequests();
  if (records.length === 0) {
    await ctx.send(buildSuccessCard("No Unactioned Requests", "There are no pending requests from this month."));
    return;
  }

  let approved = 0;
  for (const r of records) {
    const updated = await updateLeaveStatus(r.employee, r.date, "Approved", ctx.userName);
    if (updated) {
      approved++;
      const employee = await findEmployee(r.employee);
      if (employee?.teams_id) {
        await sendStatusCardToEmployee(
          nctx, employee.teams_id, "", "", r.type,
          formatDisplayDate(r.date), "Approved", ctx.userName
        );
        await sendApprovalAnnouncement(nctx, r.employee, r.type, r.date, formatDisplayDate(r.date), r.end_date);
      }
    }
  }

  await ctx.send(buildSuccessCard("Bulk Approved", `${approved} request(s) approved.`));
}

// ── approve unactioned [name] ──────────────────────────────────────────────

export async function handleApproveUnactionedForEmployee(
  ctx:  CommandContext,
  nctx: NotificationContext
): Promise<void> {
  const nameMatch = ctx.userMessage.match(/approve\s+unactioned\s+(.+)/i);
  const name      = nameMatch ? nameMatch[1].trim() : "";

  if (!name) {
    await ctx.send(buildErrorCard("Usage: `approve unactioned [employee name]`"));
    return;
  }

  const records = (await getMonthlyPendingRequests()).filter(
    (r) => r.employee.toLowerCase().includes(name.toLowerCase())
  );

  if (records.length === 0) {
    await ctx.send(buildErrorCard(`No unactioned requests found for ${name}.`));
    return;
  }

  let approved = 0;
  for (const r of records) {
    const updated = await updateLeaveStatus(r.employee, r.date, "Approved", ctx.userName);
    if (updated) {
      approved++;
      const employee = await findEmployee(r.employee);
      if (employee?.teams_id) {
        await sendStatusCardToEmployee(
          nctx, employee.teams_id, "", "", r.type,
          formatDisplayDate(r.date), "Approved", ctx.userName
        );
      }
    }
  }

  await ctx.send(buildSuccessCard("Approved", `${approved} request(s) approved for ${name}.`));
}

// ── reject all unactioned [reason] ────────────────────────────────────────

export async function handleRejectAllUnactioned(
  ctx:  CommandContext,
  nctx: NotificationContext
): Promise<void> {
  const match  = ctx.userMessage.match(/reject\s+all\s+unactioned\s+(.+)/i);
  const reason = match ? match[1].trim() : "Not actioned by approver — HR decision";

  const records = await getMonthlyPendingRequests();
  if (records.length === 0) {
    await ctx.send(buildSuccessCard("No Unactioned Requests", "There are no pending requests from this month."));
    return;
  }

  let rejected = 0;
  for (const r of records) {
    const updated = await updateLeaveStatus(r.employee, r.date, "Rejected", ctx.userName, reason);
    if (updated) {
      rejected++;
      const employee = await findEmployee(r.employee);
      if (employee?.teams_id) {
        await sendStatusCardToEmployee(
          nctx, employee.teams_id, "", "", r.type,
          formatDisplayDate(r.date), "Rejected", ctx.userName, reason
        );
      }
    }
  }

  await ctx.send(buildSuccessCard("Bulk Rejected", `${rejected} request(s) rejected. Reason: ${reason}`));
}

// ── org summary [date?] ────────────────────────────────────────────────────

export async function handleOrgSummary(ctx: CommandContext): Promise<void> {
  // Supports: "org summary", "org summary today", "org summary 2026-04-10"
  const date    = extractDate(ctx.cmd);
  const records = await getAbsencesForDateRange(date, date);
  const { buildDailySummaryCard } = await import("../cards.js");
  await ctx.send(buildDailySummaryCard(records as any[]));
}

// ── leave history [name] ───────────────────────────────────────────────────

export async function handleHRLeaveHistory(ctx: CommandContext): Promise<void> {
  const nameMatch = ctx.userMessage.match(/leave\s+history\s+(.+)/i);
  const name      = nameMatch ? nameMatch[1].trim() : "";

  if (!name) { await ctx.send(buildErrorCard("Usage: `leave history [name]`")); return; }

  const records = await getLeaveRequestsByEmployee(name);
  await ctx.send(buildMyRequestsCard(name, records as any[]));
}

// ── balance [name] ─────────────────────────────────────────────────────────

export async function handleHRBalance(ctx: CommandContext): Promise<void> {
  const nameMatch = ctx.userMessage.match(/^balance\s+(.+)/i);
  const name      = nameMatch ? nameMatch[1].trim() : "";

  if (!name) { await ctx.send(buildErrorCard("Usage: `balance [employee name]`")); return; }

  const employee = await findEmployee(name);
  if (!employee) { await ctx.send(buildErrorCard(`Employee ${name} not found.`)); return; }

  const allRecords  = await getLeaveRequestsByEmployee(name);
  const pendingDays = allRecords
    .filter((r) => r.status === "Pending" && r.type !== "WFH")
    .reduce((s, r) => s + r.days_count, 0);
  const balance = await getLeaveBalance(name);

  await ctx.send(buildLeaveBalanceCard(name, balance, pendingDays, (employee as any).carry_forward));
}

// ── view employee [name] ───────────────────────────────────────────────────

export async function handleViewEmployee(ctx: CommandContext): Promise<void> {
  const nameMatch = ctx.userMessage.match(/view\s+employee\s+(.+)/i);
  const name      = nameMatch ? nameMatch[1].trim() : "";

  if (!name) { await ctx.send(buildErrorCard("Usage: `view employee [name]`")); return; }

  const employee = await findEmployee(name);
  if (!employee) { await ctx.send(buildErrorCard(`Employee ${name} not found.`)); return; }

  await ctx.send(buildEmployeeProfileCard(employee as any));
}

// ── unregistered ───────────────────────────────────────────────────────────

export async function handleUnregistered(ctx: CommandContext): Promise<void> {
  const employees = await getUnregisteredEmployees();
  await ctx.send(buildUnregisteredCard(employees as any[]));
}

// ── org chart ──────────────────────────────────────────────────────────────

export async function handleOrgChart(ctx: CommandContext): Promise<void> {
  const all       = await getAllEmployees();
  const approvers = all.filter((e) => e.bot_role === "approver" || e.bot_role === "hr");

  const chart = approvers.map((a) => ({
    name:      a.name,
    reportees: all
      .filter((e) => e.teamlead === a.name || e.manager === a.name)
      .map((e) => e.name),
  }));

  await ctx.send(buildOrgChartCard(chart));
}

// ── team [approver name] ───────────────────────────────────────────────────

export async function handleTeamOf(ctx: CommandContext): Promise<void> {
  const nameMatch = ctx.userMessage.match(/^team\s+(.+)/i);
  const name      = nameMatch ? nameMatch[1].trim() : "";

  if (!name) { await ctx.send(buildErrorCard("Usage: `team [approver name]`")); return; }

  const all       = await getAllEmployees();
  const reportees = all
    .filter((e) => e.teamlead === name || e.manager === name)
    .map((e) => e.name);

  if (reportees.length === 0) {
    await ctx.send(buildErrorCard(`No reportees found for ${name}.`));
    return;
  }

  await ctx.send(buildSuccessCard(`Team of ${name}`, reportees.join(", ")));
}

// ── audit log ─────────────────────────────────────────────────────────────

export async function handleAuditLog(ctx: CommandContext): Promise<void> {
  const rangeMatch = ctx.userMessage.match(/(\d{4}-\d{2}-\d{2})\s+to\s+(\d{4}-\d{2}-\d{2})/i);
  let   entries    = await getAuditLog(50);
  let   label: string | undefined;

  if (rangeMatch) {
    const start = new Date(rangeMatch[1]);
    const end   = new Date(rangeMatch[2]);
    entries     = entries.filter((e) => {
      const t = new Date(e.timestamp);
      return t >= start && t <= end;
    });
    label = `${rangeMatch[1]} to ${rangeMatch[2]}`;
  }

  await ctx.send(buildAuditLogCard(entries as any[], label));
}

// ── adjust balance [name] [+/-days] [reason] ──────────────────────────────

export async function handleAdjustBalance(
  ctx:  CommandContext,
  nctx: NotificationContext
): Promise<void> {
  const match = ctx.userMessage.match(/adjust\s+balance\s+(.+?)\s+([+-]?\d+(?:\.\d+)?)\s+(.+)/i);
  if (!match) {
    await ctx.send(buildErrorCard(
      "Usage: `adjust balance [name] [+/-days] [reason]`\n" +
      "Example: `adjust balance Rithika MR +3 carry forward`"
    ));
    return;
  }

  const [, name, daysStr, reason] = match;
  const days     = parseFloat(daysStr);
  const employee = await findEmployee(name);

  if (!employee) { await ctx.send(buildErrorCard(`Employee ${name} not found.`)); return; }

  const updated = await adjustLeaveBalance(name, days, reason, ctx.userName);
  if (!updated)  { await ctx.send(buildErrorCard(`Could not update balance for ${name}.`)); return; }

  await ctx.send(buildSuccessCard(
    "Balance Updated",
    `${name}'s balance adjusted by ${days > 0 ? "+" : ""}${days} days.\nNew balance: ${updated.leave_balance} days.\nReason: ${reason}`
  ));

  if (employee.teams_id) {
    await sendBalanceAdjustedNotification(nctx, employee.teams_id, name, days, updated.leave_balance, reason, ctx.userName);
  }
}

// ── set balance [name] [days] ─────────────────────────────────────────────

export async function handleSetBalance(
  ctx:  CommandContext,
  nctx: NotificationContext
): Promise<void> {
  const match = ctx.userMessage.match(/set\s+balance\s+(.+?)\s+(\d+(?:\.\d+)?)/i);
  if (!match) {
    await ctx.send(buildErrorCard("Usage: `set balance [name] [days]`"));
    return;
  }

  const [, name, daysStr] = match;
  const targetDays = parseFloat(daysStr);
  const employee   = await findEmployee(name);

  if (!employee) { await ctx.send(buildErrorCard(`Employee ${name} not found.`)); return; }

  const diff    = targetDays - employee.leave_balance;
  const updated = await adjustLeaveBalance(name, diff, `Set to ${targetDays} by HR`, ctx.userName);

  if (!updated) { await ctx.send(buildErrorCard(`Could not set balance for ${name}.`)); return; }

  await ctx.send(buildSuccessCard("Balance Set", `${name}'s balance set to ${targetDays} days.`));

  if (employee.teams_id) {
    await sendBalanceAdjustedNotification(
      nctx, employee.teams_id, name, diff, targetDays,
      `Balance set to ${targetDays} days by HR`, ctx.userName
    );
  }
}

// ── reset balances [year] ─────────────────────────────────────────────────

export async function handleResetBalances(ctx: CommandContext): Promise<void> {
  const yearMatch = ctx.userMessage.match(/reset\s+balances\s+(\d{4})/i);
  const year      = yearMatch ? parseInt(yearMatch[1]) : new Date().getFullYear();
  const all       = await getAllEmployees();
  let   count     = 0;

  for (const emp of all) {
    const diff = 22 - emp.leave_balance;
    await adjustLeaveBalance(emp.name, diff, `Annual reset for ${year}`, ctx.userName);
    count++;
  }

  await ctx.send(buildSuccessCard(
    "Balances Reset",
    `${count} employee balances reset to 22 days for ${year}.`
  ));
}

// ── add leave for [name] [type] [date] ────────────────────────────────────

export async function handleAddLeaveOnBehalf(
  ctx:  CommandContext,
  nctx: NotificationContext
): Promise<void> {
  const match = ctx.userMessage.match(
    /add\s+leave\s+for\s+(.+?)\s+(WFH|LEAVE|SICK|MATERNITY|PATERNITY|MARRIAGE|ADOPTION)\s+(\d{4}-\d{2}-\d{2})(?:\s+to\s+(\d{4}-\d{2}-\d{2}))?/i
  );
  if (!match) {
    await ctx.send(buildErrorCard(
      "Usage: `add leave for [name] [type] [YYYY-MM-DD]`\n" +
      "Example: `add leave for Rithika MR SICK 2026-01-10`"
    ));
    return;
  }

  const [, name, type, date, endDate] = match;
  const employee = await findEmployee(name);
  if (!employee) { await ctx.send(buildErrorCard(`Employee ${name} not found.`)); return; }

  await addLeaveRequest({
    employee:    name,
    email:       employee.email,
    type:        type.toUpperCase(),
    date,
    end_date:    endDate ?? undefined,
    duration:    endDate ? "multi_day" : "full_day",
    days_count:  1,
    reason:      `Added by HR (${ctx.userName})`,
    status:      "Approved",
    approved_by: ctx.userName,
  });

  await appendAuditLog({
    hr_name:         ctx.userName,
    action:          "add_leave_behalf",
    target_employee: name,
    details:         `Added ${type} on ${date}${endDate ? ` to ${endDate}` : ""} on behalf of ${name}`,
  });

  await ctx.send(buildSuccessCard(
    "Leave Added",
    `${type} leave added for ${name} on ${formatDisplayDate(date)}. Auto-approved.`
  ));

  if (employee.teams_id) {
    await sendStatusCardToEmployee(nctx, employee.teams_id, "", "", type, formatDisplayDate(date), "Approved", ctx.userName);
    await sendApprovalAnnouncement(nctx, name, type, date, formatDisplayDate(date), endDate);
  }
}

// ── approve leave [name] [date] ────────────────────────────────────────────

export async function handleHRApproveLeave(
  ctx:  CommandContext,
  nctx: NotificationContext
): Promise<void> {
  const match = ctx.userMessage.match(/approve\s+leave\s+(.+?)\s+(\d{4}-\d{2}-\d{2})/i);
  if (!match) { await ctx.send(buildErrorCard("Usage: `approve leave [name] [YYYY-MM-DD]`")); return; }

  const [, name, date] = match;
  const updated        = await updateLeaveStatus(name, date, "Approved", ctx.userName);

  if (!updated) {
    await ctx.send(buildErrorCard(`No pending request found for ${name} on ${formatDisplayDate(date)}.`));
    return;
  }

  await ctx.send(buildSuccessCard("Approved", `${name}'s request on ${formatDisplayDate(date)} approved.`));

  const employee = await findEmployee(name);
  if (employee?.teams_id) {
    await sendStatusCardToEmployee(nctx, employee.teams_id, "", "", updated.type, formatDisplayDate(date), "Approved", ctx.userName);
    await sendApprovalAnnouncement(nctx, name, updated.type, date, formatDisplayDate(date), updated.end_date);
    await sendHRAlert(nctx, "approved", name, updated.type, formatDisplayDate(date), ctx.userName);
  }
}

// ── reject leave [name] [date] [reason] ────────────────────────────────────

export async function handleHRRejectLeave(
  ctx:  CommandContext,
  nctx: NotificationContext
): Promise<void> {
  const match = ctx.userMessage.match(/reject\s+leave\s+(.+?)\s+(\d{4}-\d{2}-\d{2})\s+(.+)/i);
  if (!match) { await ctx.send(buildErrorCard("Usage: `reject leave [name] [YYYY-MM-DD] [reason]`")); return; }

  const [, name, date, reason] = match;
  const updated                = await updateLeaveStatus(name, date, "Rejected", ctx.userName, reason);

  if (!updated) {
    await ctx.send(buildErrorCard(`No pending request found for ${name} on ${formatDisplayDate(date)}.`));
    return;
  }

  await ctx.send(buildSuccessCard("Rejected", `${name}'s request on ${formatDisplayDate(date)} rejected. Reason: ${reason}`));

  const employee = await findEmployee(name);
  if (employee?.teams_id) {
    await sendStatusCardToEmployee(nctx, employee.teams_id, "", "", updated.type, formatDisplayDate(date), "Rejected", ctx.userName, reason);
  }
}

// ── delete request [name] [date] ──────────────────────────────────────────

export async function handleHRDeleteRequest(
  ctx:  CommandContext,
  nctx: NotificationContext
): Promise<void> {
  const match = ctx.userMessage.match(/delete\s+request\s+(.+?)\s+(\d{4}-\d{2}-\d{2})/i);
  if (!match) { await ctx.send(buildErrorCard("Usage: `delete request [name] [YYYY-MM-DD]`")); return; }

  const [, name, date] = match;
  const employee       = await findEmployee(name);
  const records        = await getLeaveRequestsByEmployee(name);
  const record         = records.find((r) => r.date === date);

  if (!record) {
    await ctx.send(buildErrorCard(`No request found for ${name} on ${formatDisplayDate(date)}.`));
    return;
  }

  const reason  = "Deleted by HR";
  const deleted = await deleteLeaveRequest(name, date, ctx.userName, reason);

  if (!deleted) { await ctx.send(buildErrorCard("Could not delete the request.")); return; }

  await ctx.send(buildSuccessCard("Deleted", `${name}'s request on ${formatDisplayDate(date)} deleted.`));

  if (employee) {
    const approverTeamsId = employee.role === "teamlead"
      ? employee.manager_teams_id  ?? ""
      : employee.teamlead_teams_id ?? "";

    await sendDeleteNotifications(
      nctx, employee.teams_id ?? "", approverTeamsId,
      name, record.type, formatDisplayDate(date), ctx.userName, reason
    );
  }
}

// ── restore request [name] [date] ─────────────────────────────────────────

export async function handleRestoreRequest(ctx: CommandContext): Promise<void> {
  const match = ctx.userMessage.match(/restore\s+request\s+(.+?)\s+(\d{4}-\d{2}-\d{2})/i);
  if (!match) { await ctx.send(buildErrorCard("Usage: `restore request [name] [YYYY-MM-DD]`")); return; }

  const [, name, date] = match;

  const { PrismaClient } = await import("@prisma/client");
  const prisma = new PrismaClient();
  const record = await prisma.leaveRequest.findFirst({
    where: {
      employee: { equals: name, mode: "insensitive" },
      date,
      status: "Deleted",
    },
  });

  if (!record) {
    await ctx.send(buildErrorCard(`No deleted request found for ${name} on ${formatDisplayDate(date)}.`));
    await prisma.$disconnect();
    return;
  }

  await prisma.leaveRequest.update({
    where: { id: record.id },
    data:  { status: "Pending", deleted_by: null, deleted_at: null },
  });

  await appendAuditLog({
    hr_name:         ctx.userName,
    action:          "restore_request",
    target_employee: name,
    details:         `Restored ${record.type} on ${date}`,
  });

  await prisma.$disconnect();
  await ctx.send(buildSuccessCard(
    "Restored",
    `${name}'s ${record.type} request on ${formatDisplayDate(date)} restored to Pending.`
  ));
}

// ── add holiday [date] [name] ─────────────────────────────────────────────

export async function handleAddHoliday(
  ctx:  CommandContext,
  nctx: NotificationContext
): Promise<void> {
  const match = ctx.userMessage.match(/add\s+holiday\s+(\d{4}-\d{2}-\d{2})\s+(.+)/i);
  if (!match) { await ctx.send(buildErrorCard("Usage: `add holiday YYYY-MM-DD [Holiday Name]`")); return; }

  const [, date, name] = match;
  await addHoliday(date, name, ctx.userName);
  await ctx.send(buildSuccessCard("Holiday Added", `${name} added on ${formatDisplayDate(date)}.`));
  await sendHolidayNotificationToAll(nctx, date, name, ctx.userName, "added");
}

// ── edit holiday [date] [new name] ────────────────────────────────────────

export async function handleEditHoliday(
  ctx:  CommandContext,
  nctx: NotificationContext
): Promise<void> {
  const match = ctx.userMessage.match(/edit\s+holiday\s+(\d{4}-\d{2}-\d{2})\s+(.+)/i);
  if (!match) { await ctx.send(buildErrorCard("Usage: `edit holiday YYYY-MM-DD [New Name]`")); return; }

  const [, date, newName] = match;
  await addHoliday(date, newName, ctx.userName);
  await ctx.send(buildSuccessCard("Holiday Updated", `Holiday on ${formatDisplayDate(date)} renamed to ${newName}.`));
  await sendHolidayNotificationToAll(nctx, date, newName, ctx.userName, "edited");
}

// ── reschedule holiday [name] to [date] ───────────────────────────────────

export async function handleRescheduleHoliday(
  ctx:  CommandContext,
  nctx: NotificationContext
): Promise<void> {
  const match = ctx.userMessage.match(/reschedule\s+holiday\s+(.+?)\s+to\s+(\d{4}-\d{2}-\d{2})/i);
  if (!match) {
    await ctx.send(buildErrorCard("Usage: `reschedule holiday [Holiday Name] to YYYY-MM-DD`"));
    return;
  }

  const [, holidayName, newDate] = match;

  const { PrismaClient } = await import("@prisma/client");
  const prisma = new PrismaClient();
  const existing = await prisma.holiday.findFirst({
    where: { name: { contains: holidayName, mode: "insensitive" } },
  });

  if (!existing) {
    await ctx.send(buildErrorCard(`No holiday found named "${holidayName}".`));
    await prisma.$disconnect();
    return;
  }

  await prisma.holiday.delete({ where: { id: existing.id } });
  await prisma.$disconnect();

  await addHoliday(newDate, existing.name, ctx.userName);
  await appendAuditLog({
    hr_name:         ctx.userName,
    action:          "reschedule_holiday",
    target_employee: null,
    details:         `Rescheduled ${existing.name} from ${existing.date} to ${newDate}`,
  });

  await ctx.send(buildSuccessCard("Holiday Rescheduled", `${existing.name} moved to ${formatDisplayDate(newDate)}.`));
  await sendHolidayNotificationToAll(nctx, newDate, existing.name, ctx.userName, "rescheduled");
}

// ── delete holiday [date] ─────────────────────────────────────────────────

export async function handleDeleteHoliday(
  ctx:  CommandContext,
  nctx: NotificationContext
): Promise<void> {
  const match = ctx.userMessage.match(/delete\s+holiday\s+(\d{4}-\d{2}-\d{2})/i);
  if (!match) { await ctx.send(buildErrorCard("Usage: `delete holiday YYYY-MM-DD`")); return; }

  const date = match[1];
  const { PrismaClient } = await import("@prisma/client");
  const prisma   = new PrismaClient();
  const existing = await prisma.holiday.findUnique({ where: { date } });

  if (!existing) {
    await ctx.send(buildErrorCard(`No holiday found on ${formatDisplayDate(date)}.`));
    await prisma.$disconnect();
    return;
  }

  await prisma.holiday.delete({ where: { date } });
  await prisma.$disconnect();

  await appendAuditLog({
    hr_name:         ctx.userName,
    action:          "delete_holiday",
    target_employee: null,
    details:         `Deleted holiday: ${existing.name} on ${date}`,
  });

  await ctx.send(buildSuccessCard("Holiday Removed", `${existing.name} on ${formatDisplayDate(date)} removed.`));
  await sendHolidayNotificationToAll(nctx, date, existing.name, ctx.userName, "deleted");
}

export async function handleAddMultipleHolidays(ctx: CommandContext): Promise<void> {
  const match = ctx.userMessage.match(/add\s+multiple\s+holidays\s+/i);
  if (!match){await ctx.send(buildErrorCard("Usage:\nadd multiple holidays\nYYYY-MM-DD Holiday Name\n...")); return; }

  
  const lines = ctx.userMessage
    .replace(/^add multiple holidays/i, "")
    .split("\n")
    .map((l) => l.trim())
    .filter(Boolean);

  if (lines.length === 0) {
    await ctx.send(buildErrorCard(
      "Usage:\nadd multiple holidays\nYYYY-MM-DD Holiday Name\n..."
    ));
    return;
  }

  let added = 0;
  let updated = 0;
  let skipped = 0;

  for (const line of lines) {
    const match = line.match(/^(\d{4}-\d{2}-\d{2})\s+(.+)$/);

    if (!match) {
      skipped++;
      continue;
    }

    const [, date, name] = match;

    try {
      const existing = await isHoliday(date);

      if (existing) {
        // ✅ Override (update name)
        //await updateHoliday(date, name, ctx.userName);
        updated++;
      } else {
        await addHoliday(date, name, ctx.userName);
        added++;
      }

    } catch {
      skipped++;
    }
  }

  await ctx.send(buildSuccessCard(
    "Holidays Updated",
    `Added: ${added}\nUpdated: ${updated}\nSkipped: ${skipped}`
  ));
}

// ── Clear All Holidays ──────────────────────────────────────────

export async function handleClearHolidays(ctx: CommandContext): Promise<void> {
  

  const match = ctx.userMessage.match(/clear\s+all\s+holidays\s+/i);
  if (!match){await ctx.send(buildErrorCard("Usage:\nclear all holidays")); return; }

  const count = await clearAllHolidays(ctx.userName);

  await ctx.send(
    count === 0
      ? "No holidays found to clear."
      : `All holidays cleared successfully (${count} removed).`
  );
}


// ── add employee ───────────────────────────────────────────────────────────
// Usage: add employee name: X | email: X | role: X | bot_role: employee | teamlead: X | manager: X | leave_balance: 22

export async function handleAddEmployee(ctx: CommandContext): Promise<void> {
  // Parse pipe-separated key: value fields (case-insensitive keys)
  const body   = ctx.userMessage.replace(/^add\s+employee\s*/i, "");
  const fields: Record<string, string> = {};

  for (const part of body.split("|")) {
    const colonIdx = part.indexOf(":");
    if (colonIdx === -1) continue;
    const key = part.slice(0, colonIdx).trim().toLowerCase().replace(/\s+/g, "_");
    const val = part.slice(colonIdx + 1).trim();
    if (key && val) fields[key] = val;
  }

  const required = ["name", "email", "role", "bot_role"];
  const missing  = required.filter((r) => !fields[r]);
  if (missing.length) {
    await ctx.send(buildErrorCard(
      `Missing required field(s): ${missing.join(", ")}\n\n` +
      `Usage:\nadd employee name: Priya Sharma | email: priya@company.com | role: Software Engineer | bot_role: employee | teamlead: Rithika MR | manager: Anil Kumar | leave_balance: 22\n\n` +
      `bot_role options: employee, approver, hr`
    ));
    return;
  }

  // Validate bot_role
  const validRoles = ["employee", "approver", "hr"];
  if (!validRoles.includes(fields.bot_role.toLowerCase())) {
    await ctx.send(buildErrorCard(`Invalid bot_role "${fields.bot_role}". Must be one of: employee, approver, hr`));
    return;
  }

  // Check for duplicate
  const existing = await findEmployee(fields.name);
  if (existing) {
    await ctx.send(buildErrorCard(
      `Employee "${fields.name}" already exists.\nUse: update employee ${fields.name} [field] [value]`
    ));
    return;
  }

  // Resolve teams_ids for reporting lines
  const allEmps = await getAllEmployees();
  const tl  = allEmps.find((e) => e.name.toLowerCase() === fields.teamlead?.toLowerCase());
  const mgr = allEmps.find((e) => e.name.toLowerCase() === fields.manager?.toLowerCase());

await upsertEmployee({
  name:              fields.name,
  email:             fields.email,
  role:              fields.role,
  bot_role:          fields.bot_role.toLowerCase() as "employee" | "approver" | "hr",

  teamlead:          fields.teamlead ?? null,
  manager:           fields.manager ?? null,

  teamlead_email:    tl?.email ?? null,
  manager_email:     mgr?.email ?? null,

  teamlead_teams_id: tl?.teams_id ?? null,
  manager_teams_id:  mgr?.teams_id ?? null,

  leave_balance:     parseFloat(fields.leave_balance ?? "22"),
  carry_forward:     parseFloat(fields.carry_forward ?? "0"),

  year_entitlement_start: new Date().getFullYear(),  teams_id:          fields.teams_id ?? null,
});

  await appendAuditLog({
    hr_name:         ctx.userName,
    action:          "add_employee",
    target_employee: fields.name,
    details:         `Added ${fields.name} as ${fields.bot_role} (${fields.role})`,
  });

  await ctx.send(buildSuccessCard(
    "Employee Added",
    `${fields.name} added successfully.\n\n` +
    `Role: ${fields.role} | Bot role: ${fields.bot_role}\n` +
    `Balance: ${fields.leave_balance ?? 22} days\n\n` +
    `Note: teams_id will auto-register when ${fields.name} sends their first message to the bot.`
  ));
}

// ── update employee [name] [field] [value] ────────────────────────────────
// Usage: update employee Rithika MR bot_role approver
// Fields: name, email, role, bot_role, manager, teamlead, teams_id,
//         leave_balance, carry_forward

export async function handleUpdateEmployee(ctx: CommandContext): Promise<void> {
  const UPDATABLE_FIELDS = [
    "name", "email", "role", "bot_role",
    "manager", "teamlead", "teams_id",
    "leave_balance", "carry_forward",
  ];

  // Build a regex that matches any known field name (case-insensitive)
  const fieldPattern = UPDATABLE_FIELDS.join("|");
  const match = ctx.userMessage.match(
    new RegExp(`update\\s+employee\\s+(.+?)\\s+(${fieldPattern})\\s+(.+)`, "i")
  );

  if (!match) {
    await ctx.send(buildErrorCard(
      "Usage: `update employee [name] [field] [value]`\n\n" +
      `Fields: ${UPDATABLE_FIELDS.join(", ")}\n\n` +
      "Examples:\n" +
      "  update employee Priya Sharma bot_role approver\n" +
      "  update employee Priya Sharma role Senior Engineer\n" +
      "  update employee Priya Sharma manager Anil Kumar\n" +
      "  update employee Priya Sharma leave_balance 18"
    ));
    return;
  }

  const [, name, field, value] = match;
  const fieldLower = field.toLowerCase();

  const employee = await findEmployee(name);
  if (!employee) { await ctx.send(buildErrorCard(`Employee "${name}" not found.`)); return; }

  // Validate bot_role changes
  if (fieldLower === "bot_role") {
    const validRoles = ["employee", "approver", "hr"];
    if (!validRoles.includes(value.toLowerCase())) {
      await ctx.send(buildErrorCard(`Invalid bot_role "${value}". Must be one of: employee, approver, hr`));
      return;
    }
  }

  // Build the update payload
  const allEmps = await getAllEmployees();
  const updates: Record<string, any> = { [fieldLower]: value };

  // When changing teamlead/manager, also update the corresponding teams_id
  if (fieldLower === "teamlead") {
    const tl = allEmps.find((e) => e.name.toLowerCase() === value.toLowerCase());
    updates.teamlead_teams_id = tl?.teams_id ?? null;
  }
  if (fieldLower === "manager") {
    const mgr = allEmps.find((e) => e.name.toLowerCase() === value.toLowerCase());
    updates.manager_teams_id = mgr?.teams_id ?? null;
  }

  // Coerce numeric fields
  if (fieldLower === "leave_balance" || fieldLower === "carry_forward") {
    updates[fieldLower] = parseFloat(value);
  }

  await upsertEmployee({ ...employee, ...updates });

  await appendAuditLog({
    hr_name:         ctx.userName,
    action:          "update_employee",
    target_employee: name,
    details:         `Updated ${fieldLower} → "${value}"`,
  });

  // Extra note for reporting-line changes: update affected reportees
  let note = "";
  if (fieldLower === "bot_role" || fieldLower === "teamlead" || fieldLower === "manager") {
    note = `\n\nRemember: if other employees report to ${name}, update their teamlead/manager fields too.`;
  }

  await ctx.send(buildSuccessCard(
    "Employee Updated",
    `${name} — ${fieldLower} updated to "${value}".${note}`
  ));
}

// ── deactivate employee [name] ────────────────────────────────────────────

export async function handleDeactivateEmployee(ctx: CommandContext): Promise<void> {
  const nameMatch = ctx.userMessage.match(/deactivate\s+employee\s+(.+)/i);
  const name      = nameMatch ? nameMatch[1].trim() : "";

  if (!name) { await ctx.send(buildErrorCard("Usage: `deactivate employee [name]`")); return; }

  const employee = await findEmployee(name);
  if (!employee) { await ctx.send(buildErrorCard(`Employee "${name}" not found.`)); return; }

  if (!(employee as any).is_active) {
    await ctx.send(buildErrorCard(`${name} is already inactive.`));
    return;
  }

  await upsertEmployee({ ...employee, is_active: false } as any);

  await appendAuditLog({
    hr_name:         ctx.userName,
    action:          "deactivate_employee",
    target_employee: name,
    details:         `Deactivated ${name} — marked inactive`,
  });

  await ctx.send(buildSuccessCard(
    "Employee Deactivated",
    `${name} has been marked inactive.\n\n` +
    `They will no longer appear in org chart, team lists, or receive notifications.\n` +
    `Their historical leave records are preserved.\n\n` +
    `To reactivate: update employee ${name} is_active true`
  ));
}

// ── download report [month] [year] ────────────────────────────────────────

export async function handleDownloadReport(ctx: CommandContext): Promise<void> {
  const monthInfo = extractMonth(ctx.cmd);
  const yearMatch = ctx.userMessage.match(/\b(20\d{2})\b/);
  const isYtd     = /ytd|year.to.date/i.test(ctx.cmd);

  const year  = yearMatch ? parseInt(yearMatch[1]) : new Date().getFullYear();
  const month = monthInfo ? parseInt(monthInfo.mm) : new Date().getMonth() + 1;
  const mm    = String(month).padStart(2, "0");
  const yyyy  = String(year);

  const all  = await getAllLeaveRequests();
  const emps = await getAllEmployees();

  const filtered = isYtd
    ? all.filter((r) => r.date.startsWith(yyyy))
    : all.filter((r) => r.date.startsWith(`${yyyy}-${mm}`));

  const rows = emps.map((emp) => {
    const empRecords = filtered.filter((r) => r.employee === emp.name);
    const leaveTaken = empRecords
      .filter((r) => r.type !== "WFH" && r.status === "Approved")
      .reduce((s, r) => s + r.days_count, 0);
    const wfhDays = empRecords
      .filter((r) => r.type === "WFH" && r.status === "Approved")
      .reduce((s, r) => s + r.days_count, 0);
    const pending = empRecords
      .filter((r) => r.status === "Pending")
      .reduce((s, r) => s + r.days_count, 0);
    const closing = Math.max(0, emp.leave_balance);
    const opening = closing + leaveTaken;

    return {
      "Employee Name":   emp.name,
      "Month":           isYtd ? yyyy : `${yyyy}-${mm}`,
      "Opening Balance": opening,
      "Leave Taken":     leaveTaken,
      "WFH Days":        wfhDays,
      "Pending":         pending,
      "Closing Balance": closing,
    };
  });

  const wb         = XLSX.utils.book_new();
  const ws         = XLSX.utils.json_to_sheet(rows);
  XLSX.utils.book_append_sheet(wb, ws, "Report");

  const reportName = isYtd
    ? `LeaveReport_YTD_${yyyy}.xlsx`
    : `LeaveReport_${yyyy}_${mm}.xlsx`;
  const outPath    = path.join(process.cwd(), "data", reportName);
  XLSX.writeFile(wb, outPath);

  await ctx.send(buildSuccessCard(
    "Report Generated",
    `File saved: data/${reportName}\nOpen from the data folder on the server.`
  ));
}

// ── remind approvers ───────────────────────────────────────────────────────

export async function handleRemindApprovers(
  ctx:  CommandContext,
  nctx: NotificationContext
): Promise<void> {
  const records    = await getMonthlyPendingRequests();
  const now        = new Date();
  const monthLabel = now.toLocaleDateString("en-IN", { month: "long", year: "numeric" });

  const allEmps = await getAllEmployees();
  const approverGroups: Record<string, any[]> = {};

  for (const r of records) {
    const emp = allEmps.find((e) => e.name.toLowerCase() === r.employee.toLowerCase());
    if (!emp) continue;
    const approverTeamsId = emp.role === "teamlead" ? emp.manager_teams_id  : emp.teamlead_teams_id;
    const approverName    = emp.role === "teamlead" ? emp.manager           : emp.teamlead;
    if (!approverTeamsId || !approverName) continue;
    if (!approverGroups[approverTeamsId]) approverGroups[approverTeamsId] = [];
    approverGroups[approverTeamsId].push({ approverName, approverTeamsId, record: r });
  }

  const { sendApproverReminders } = await import("../notificationServices.js");
  const groups = Object.entries(approverGroups).map(([tid, items]) => ({
    approverName:    items[0].approverName,
    approverTeamsId: tid,
    records:         items.map((i) => i.record),
  }));

  await sendApproverReminders(groups, monthLabel);
  await ctx.send(buildSuccessCard(
    "Reminders Sent",
    `Reminders sent to ${groups.length} approver(s) with pending requests.`
  ));
}

// ── who is on leave / wfh — full org (HR) ────────────────────────────────

export async function handleHRWhoIsOnLeave(ctx: CommandContext): Promise<void> {
  const ctxWithOverride = { ...ctx, role: { ...ctx.role, botRole: "hr" as const } };
  await handleWhoIsOnLeave(ctxWithOverride, false);
}

export async function handleHRWhoIsOnLeaveImpl(ctx: CommandContext): Promise<void> {
  await handleHRWhoIsOnLeave(ctx);
}