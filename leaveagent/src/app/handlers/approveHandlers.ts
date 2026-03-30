import {
  buildTeamRequestsCard,
  buildPendingApprovalsCard,
  buildWhoIsOnLeaveCard,
  buildLeaveBalanceCard,
  buildMyRequestsCard,
  buildSuccessCard,
  buildErrorCard,
  formatDisplayDate,
} from "../cards";
import {
  getLeaveRequestsForApprover,
  getPendingRequestsForApprover,
  getAbsencesForDateRange,
  findEmployee,
  getLeaveRequestsByEmployee,
  updateLeaveStatus,
  getLeaveBalance,
} from "../postgresManager";
import { sendStatusCardToEmployee, sendApprovalAnnouncement, sendWorkforceCardToManager, sendHRAlert, NotificationContext } from "../notificationServices";
import { CommandContext } from "./sharedHandlers";

// ── Parse date range from natural language ────────────────────────────────

function parseDateRange(cmd: string): { start: string; end: string; label: string } {
  const today      = new Date();
  const todayStr   = today.toISOString().split("T")[0];

  // "today"
  if (/today/i.test(cmd)) {
    return { start: todayStr, end: todayStr, label: "Today" };
  }

  // "tomorrow"
  if (/tomorrow/i.test(cmd)) {
    const d = new Date(today); d.setDate(d.getDate() + 1);
    const s = d.toISOString().split("T")[0];
    return { start: s, end: s, label: "Tomorrow" };
  }

  // "next week"
  if (/next week/i.test(cmd)) {
    const d    = new Date(today);
    const day  = d.getDay();
    const diff = day === 0 ? 1 : 8 - day;
    d.setDate(d.getDate() + diff);
    const start = d.toISOString().split("T")[0];
    const end   = new Date(d); end.setDate(d.getDate() + 4);
    return { start, end: end.toISOString().split("T")[0], label: "Next Week" };
  }

  // "this week"
  if (/this week/i.test(cmd)) {
    const d    = new Date(today);
    const day  = d.getDay();
    const diff = day === 0 ? 0 : 1 - day;
    d.setDate(d.getDate() + diff);
    const start = d.toISOString().split("T")[0];
    const end   = new Date(d); end.setDate(d.getDate() + 4);
    return { start, end: end.toISOString().split("T")[0], label: "This Week" };
  }

  // explicit dates: YYYY-MM-DD to YYYY-MM-DD
  const rangeMatch = cmd.match(/(\d{4}-\d{2}-\d{2})\s+to\s+(\d{4}-\d{2}-\d{2})/);
  if (rangeMatch) {
    return { start: rangeMatch[1], end: rangeMatch[2], label: `${formatDisplayDate(rangeMatch[1])} to ${formatDisplayDate(rangeMatch[2])}` };
  }

  // single date
  const dateMatch = cmd.match(/(\d{4}-\d{2}-\d{2})/);
  if (dateMatch) {
    return { start: dateMatch[1], end: dateMatch[1], label: formatDisplayDate(dateMatch[1]) };
  }

  // default: today
  return { start: todayStr, end: todayStr, label: "Today" };
}

// ── team summary ──────────────────────────────────────────────────────────

export async function handleTeamSummary(ctx: CommandContext): Promise<void> {
  const { start, end, label } = parseDateRange(ctx.cmd);
  const all     = await getLeaveRequestsForApprover(ctx.userName);
  const records = all.filter((r) => r.status === "Approved" && r.date >= start && r.date <= end);
  await ctx.send(buildTeamRequestsCard(ctx.userName, records as any[], `👥 Team Availability — ${label}`));
}

// ── team requests ──────────────────────────────────────────────────────────

export async function handleTeamRequests(ctx: CommandContext): Promise<void> {
  const monthNames = ["january","february","march","april","may","june","july","august","september","october","november","december"];
  let   records    = await getLeaveRequestsForApprover(ctx.userName);

  for (let i = 0; i < monthNames.length; i++) {
    if (ctx.cmd.includes(monthNames[i])) {
      const mm = String(i + 1).padStart(2, "0");
      const yy = String(new Date().getFullYear());
      records   = records.filter((r) => r.date.startsWith(`${yy}-${mm}`));
      break;
    }
  }

  await ctx.send(buildTeamRequestsCard(ctx.userName, records as any[]));
}

// ── pending approvals ──────────────────────────────────────────────────────

export async function handlePendingApprovals(ctx: CommandContext): Promise<void> {
  const records = await getPendingRequestsForApprover(ctx.userName);
  await ctx.send(buildPendingApprovalsCard(ctx.userName, records as any[]));
}

// ── who is on leave / who is wfh ──────────────────────────────────────────

export async function handleWhoIsOnLeave(ctx: CommandContext, scopeToTeam = true): Promise<void> {
  const { start, end, label } = parseDateRange(ctx.cmd);
  const isWfh = /who is wfh|who('s| is) wfh|wfh/i.test(ctx.cmd);

  let records = await getAbsencesForDateRange(start, end);

  // Approvers only see their team
  if (scopeToTeam && ctx.role.botRole === "approver") {
    const teamRecords = await getLeaveRequestsForApprover(ctx.userName);
    const teamNames   = new Set(teamRecords.map((r) => r.employee.toLowerCase()));
    records = records.filter((r) => teamNames.has(r.employee.toLowerCase()));
  }

  const filtered = isWfh
    ? records.filter((r) => r.type === "WFH")
    : records;

  const type = isWfh ? "wfh" : "all";
  await ctx.send(buildWhoIsOnLeaveCard(filtered as any[], label, type));
}

// ── who is available ───────────────────────────────────────────────────────

export async function handleWhoIsAvailable(ctx: CommandContext): Promise<void> {
  const { start, end, label } = parseDateRange(ctx.cmd);
  let absences = await getAbsencesForDateRange(start, end);

  if (ctx.role.botRole === "approver") {
    const teamRecords = await getLeaveRequestsForApprover(ctx.userName);
    const teamNames   = new Set(teamRecords.map((r) => r.employee.toLowerCase()));
    absences = absences.filter((r) => teamNames.has(r.employee.toLowerCase()));
  }

  const absentNames = new Set(absences.map((r) => r.employee.toLowerCase()));

  // Get all team members and filter out absent ones
  const allTeam = await getLeaveRequestsForApprover(ctx.userName);
  const teamSet = [...new Set(allTeam.map((r) => r.employee))];
  const available = teamSet.filter((name) => !absentNames.has(name.toLowerCase()));

  const fakeRecords = available.map((name) => ({
    employee: name, type: "AVAILABLE", date: start, end_date: null,
    duration: "full_day", days_count: 0, status: "Available",
  }));

  await ctx.send(buildWhoIsOnLeaveCard([], label, "all")); // will show "everyone available" if empty
  if (available.length > 0) {
    await ctx.send(`✅ Available on ${label}: ${available.join(", ")}`);
  }
}

// ── leave history [name] ───────────────────────────────────────────────────

export async function handleLeaveHistory(ctx: CommandContext): Promise<void> {
  const parts  = ctx.userMessage.split(/\s+/);
  const nameIdx = parts.findIndex((p) => /history/i.test(p)) + 1;
  const name    = parts.slice(nameIdx).join(" ").trim();

  if (!name) {
    await ctx.send(buildErrorCard("Please provide a name: `leave history Rithika MR`"));
    return;
  }

  const records = await getLeaveRequestsByEmployee(name);
  if (records.length === 0) {
    await ctx.send(buildErrorCard(`No requests found for ${name}.`));
    return;
  }

  await ctx.send(buildMyRequestsCard(name, records as any[]));
}

// ── balance [name] ─────────────────────────────────────────────────────────

export async function handleBalanceForReportee(ctx: CommandContext): Promise<void> {
  const parts   = ctx.userMessage.split(/\s+/);
  const nameIdx = parts.findIndex((p) => /balance/i.test(p)) + 1;
  const name    = parts.slice(nameIdx).join(" ").trim();

  if (!name) {
    await ctx.send(buildErrorCard("Please provide a name: `balance Rithika MR`"));
    return;
  }

  const employee = await findEmployee(name);
  if (!employee) {
    await ctx.send(buildErrorCard(`Employee ${name} not found.`));
    return;
  }

  const allRecords  = await getLeaveRequestsByEmployee(name);
  const pendingDays = allRecords
    .filter((r) => r.status === "Pending" && r.type !== "WFH")
    .reduce((sum, r) => sum + r.days_count, 0);

  const balance = await getLeaveBalance(name);
  await ctx.send(buildLeaveBalanceCard(name, balance, pendingDays, employee.carry_forward));
}

// ── approve leave [name] [date] ────────────────────────────────────────────

export async function handleApproveLeaveCommand(
  ctx:  CommandContext,
  nctx: NotificationContext
): Promise<void> {
  const match = ctx.userMessage.match(/approve leave\s+(.+?)\s+(\d{4}-\d{2}-\d{2})/i);
  if (!match) {
    await ctx.send(buildErrorCard("Usage: `approve leave [name] [YYYY-MM-DD]`"));
    return;
  }

  const [, employeeName, date] = match;
  const displayDate = formatDisplayDate(date);
  const updated     = await updateLeaveStatus(employeeName, date, "Approved", ctx.userName);

  if (!updated) {
    await ctx.send(buildErrorCard(`No pending request found for ${employeeName} on ${displayDate}.`));
    return;
  }

  await ctx.send(buildSuccessCard("Request Approved", `${employeeName}'s request for ${displayDate} has been approved.`));

  const employee = await findEmployee(employeeName);
  if (employee?.teams_id) {
    await sendStatusCardToEmployee(
      nctx, employee.teams_id, ctx.activity.from.id,
      ctx.activity.conversation.id, updated.type, displayDate,
      "Approved", ctx.userName, undefined
    );
    await sendApprovalAnnouncement(nctx, employeeName, updated.type, date, displayDate, updated.end_date);
    await sendHRAlert(nctx, "approved", employeeName, updated.type, displayDate, ctx.userName);
  }
}

// ── reject leave [name] [date] [reason] ────────────────────────────────────

export async function handleRejectLeaveCommand(
  ctx:  CommandContext,
  nctx: NotificationContext
): Promise<void> {
  const match = ctx.userMessage.match(/reject leave\s+(.+?)\s+(\d{4}-\d{2}-\d{2})\s+(.+)/i);
  if (!match) {
    await ctx.send(buildErrorCard("Usage: `reject leave [name] [YYYY-MM-DD] [reason]`"));
    return;
  }

  const [, employeeName, date, reason] = match;
  const displayDate = formatDisplayDate(date);
  const updated     = await updateLeaveStatus(employeeName, date, "Rejected", ctx.userName, reason);

  if (!updated) {
    await ctx.send(buildErrorCard(`No pending request found for ${employeeName} on ${displayDate}.`));
    return;
  }

  await ctx.send(buildSuccessCard("Request Rejected", `${employeeName}'s request for ${displayDate} has been rejected. Reason: ${reason}`));

  const employee = await findEmployee(employeeName);
  if (employee?.teams_id) {
    await sendStatusCardToEmployee(
      nctx, employee.teams_id, ctx.activity.from.id,
      ctx.activity.conversation.id, updated.type, displayDate,
      "Rejected", ctx.userName, reason
    );
  }
}
