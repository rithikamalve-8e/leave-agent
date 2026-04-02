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
  getAllEmployees,
} from "../postgresManager";
import {
  sendStatusCardToEmployee,
  sendApprovalAnnouncement,
  sendHRAlert,
  sendWorkforceCardToManager,
  NotificationContext,
} from "../notificationServices";
import { CommandContext } from "./sharedHandlers";

// ── Month name table (lowercase) ──────────────────────────────────────────

const MONTH_NAMES = [
  "january","february","march","april","may","june",
  "july","august","september","october","november","december",
];

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

// ── Parse date range from command ─────────────────────────────────────────

export function parseDateRange(cmd: string): { start: string; end: string; label: string } {
  const lower    = cmd.toLowerCase();
  const today    = new Date();
  const todayStr = today.toISOString().split("T")[0];

  // "today"
  if (/\btoday\b/.test(lower)) {
    return { start: todayStr, end: todayStr, label: "Today" };
  }

  // "tomorrow"
  if (/\btomorrow\b/.test(lower)) {
    const d = new Date(today);
    d.setDate(d.getDate() + 1);
    const s = d.toISOString().split("T")[0];
    return { start: s, end: s, label: "Tomorrow" };
  }

  // "next week"
  if (/\bnext\s+week\b/.test(lower)) {
    const d   = new Date(today);
    const day = d.getDay();                   // 0=Sun
    const diff = day === 0 ? 1 : 8 - day;     // days until next Monday
    d.setDate(d.getDate() + diff);
    const start = d.toISOString().split("T")[0];
    const end   = new Date(d);
    end.setDate(d.getDate() + 4);             // Friday of that week
    return { start, end: end.toISOString().split("T")[0], label: "Next Week" };
  }

  // "this week"
  if (/\bthis\s+week\b/.test(lower)) {
    const d   = new Date(today);
    const day = d.getDay();
    const diff = day === 0 ? -6 : 1 - day;    // back to Monday
    d.setDate(d.getDate() + diff);
    const start = d.toISOString().split("T")[0];
    const end   = new Date(d);
    end.setDate(d.getDate() + 4);
    return { start, end: end.toISOString().split("T")[0], label: "This Week" };
  }

  // "this month" or month name
  if (/\bthis\s+month\b/.test(lower)) {
    const mm    = String(today.getMonth() + 1).padStart(2, "0");
    const yyyy  = String(today.getFullYear());
    const last  = new Date(today.getFullYear(), today.getMonth() + 1, 0);
    return {
      start: `${yyyy}-${mm}-01`,
      end:   last.toISOString().split("T")[0],
      label: today.toLocaleDateString("en-IN", { month: "long", year: "numeric" }),
    };
  }

  // Named month (e.g. "march", "April")
  const monthInfo = extractMonth(cmd);
  if (monthInfo) {
    const yyyy  = String(today.getFullYear());
    const first = `${yyyy}-${monthInfo.mm}-01`;
    const last  = new Date(parseInt(yyyy), parseInt(monthInfo.mm), 0);
    return {
      start: first,
      end:   last.toISOString().split("T")[0],
      label: `${monthInfo.name} ${yyyy}`,
    };
  }

  // Explicit range: YYYY-MM-DD to YYYY-MM-DD
  const rangeMatch = cmd.match(/(\d{4}-\d{2}-\d{2})\s+to\s+(\d{4}-\d{2}-\d{2})/);
  if (rangeMatch) {
    return {
      start: rangeMatch[1],
      end:   rangeMatch[2],
      label: `${formatDisplayDate(rangeMatch[1])} to ${formatDisplayDate(rangeMatch[2])}`,
    };
  }

  // Single explicit date
  const dateMatch = cmd.match(/(\d{4}-\d{2}-\d{2})/);
  if (dateMatch) {
    return {
      start: dateMatch[1],
      end:   dateMatch[1],
      label: formatDisplayDate(dateMatch[1]),
    };
  }

  // Default: today
  return { start: todayStr, end: todayStr, label: "Today" };
}

// ── team summary ──────────────────────────────────────────────────────────

export async function handleTeamSummary(ctx: CommandContext): Promise<void> {
  const { start, end, label } = parseDateRange(ctx.cmd);
  const all     = await getLeaveRequestsForApprover(ctx.userName);
  const records = all.filter(
    (r) => r.status === "Approved" && r.date >= start && r.date <= end
  );
  await ctx.send(buildTeamRequestsCard(ctx.userName, records as any[], `Team Availability — ${label}`));
}

// ── team requests ──────────────────────────────────────────────────────────

export async function handleTeamRequests(ctx: CommandContext): Promise<void> {
  let records = await getLeaveRequestsForApprover(ctx.userName);

  // Filter by month name if present
  const monthInfo = extractMonth(ctx.cmd);
  if (monthInfo) {
    const yy = String(new Date().getFullYear());
    records  = records.filter((r) => r.date.startsWith(`${yy}-${monthInfo.mm}`));
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

  // Detect WFH-specific query (case-insensitive)
  const isWfh = /who\s+(is\s+)?wfh|wfh\s+today|wfh\s+tomorrow/i.test(ctx.cmd);

  let records = await getAbsencesForDateRange(start, end);

  // Approvers see only their team; HR sees everyone (scopeToTeam=false)
  if (scopeToTeam && ctx.role.botRole === "approver") {
    const teamRecords = await getLeaveRequestsForApprover(ctx.userName);
    const teamNames   = new Set(teamRecords.map((r) => r.employee.toLowerCase()));
    records = records.filter((r) => teamNames.has(r.employee.toLowerCase()));
  }

  const filtered = isWfh
    ? records.filter((r) => r.type.toUpperCase() === "WFH")
    : records;

  const type = isWfh ? "wfh" : "all";
  await ctx.send(buildWhoIsOnLeaveCard(filtered as any[], label, type));
}

// ── who is available ───────────────────────────────────────────────────────

export async function handleWhoIsAvailable(ctx: CommandContext): Promise<void> {
  const { start, end, label } = parseDateRange(ctx.cmd);

  // Step 1: Get absences
  let absences = await getAbsencesForDateRange(start, end);

  // Step 2: Get all employees (source of truth)
  const allEmployees = await getAllEmployees();

  let teamSet: string[];

  // Step 3: Filter team for approver
  if (ctx.role.botRole === "approver") {
    const approver = ctx.userName.toLowerCase();

    teamSet = allEmployees
      .filter((e) =>
        e.manager?.toLowerCase() === approver ||
        e.teamlead?.toLowerCase() === approver
      )
      .map((e) => e.name);

    const teamNames = new Set(teamSet.map((n) => n.toLowerCase()));

    // Filter absences only for this team
    absences = absences.filter((r) =>
      teamNames.has(r.employee.toLowerCase())
    );
  } else {
    // HR sees entire org
    teamSet = allEmployees.map((e) => e.name);
  }

  // Step 4: Compute available = team - absences
  const absentNames = new Set(
    absences.map((r) => r.employee.toLowerCase())
  );

  const available = teamSet.filter(
    (name) => !absentNames.has(name.toLowerCase())
  );

  // Step 5: Send response
  await ctx.send(
    buildSuccessCard(
      `Available on ${label}`,
      available.length
        ? available.join(", ")
        : "Everyone is available!"
    )
  );
}

// ── leave history [name] ───────────────────────────────────────────────────

export async function handleLeaveHistory(ctx: CommandContext): Promise<void> {
  const nameMatch = ctx.userMessage.match(/leave\s+history\s+(.+)/i);
  const name      = nameMatch ? nameMatch[1].trim() : "";

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
  const nameMatch = ctx.userMessage.match(/^balance\s+(.+)/i);
  const name      = nameMatch ? nameMatch[1].trim() : "";

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
  const match = ctx.userMessage.match(/approve\s+leave\s+(.+?)\s+(\d{4}-\d{2}-\d{2})/i);
  if (!match) {
    await ctx.send(buildErrorCard("Usage: `approve leave [name] [YYYY-MM-DD]`"));
    return;
  }

  const [, employeeName, date] = match;
  const displayDate = formatDisplayDate(date);
  const updated     = await updateLeaveStatus(employeeName, date, "Approved", ctx.userName);

  if (!updated) {
    await ctx.send(buildErrorCard(
      `No pending request found for ${employeeName} on ${displayDate}.`
    ));
    return;
  }

  await ctx.send(buildSuccessCard(
    "Request Approved",
    `${employeeName}'s request for ${displayDate} has been approved.`
  ));

  const employee = await findEmployee(employeeName);
  if (employee?.teams_id) {
    await sendStatusCardToEmployee(
      nctx, employee.teams_id,
      ctx.activity.from.id, ctx.activity.conversation.id,
      updated.type, displayDate, "Approved", ctx.userName, undefined
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
  const match = ctx.userMessage.match(/reject\s+leave\s+(.+?)\s+(\d{4}-\d{2}-\d{2})\s+(.+)/i);
  if (!match) {
    await ctx.send(buildErrorCard("Usage: `reject leave [name] [YYYY-MM-DD] [reason]`"));
    return;
  }

  const [, employeeName, date, reason] = match;
  const displayDate = formatDisplayDate(date);
  const updated     = await updateLeaveStatus(employeeName, date, "Rejected", ctx.userName, reason);

  if (!updated) {
    await ctx.send(buildErrorCard(
      `No pending request found for ${employeeName} on ${displayDate}.`
    ));
    return;
  }

  await ctx.send(buildSuccessCard(
    "Request Rejected",
    `${employeeName}'s request for ${displayDate} has been rejected. Reason: ${reason}`
  ));

  const employee = await findEmployee(employeeName);
  if (employee?.teams_id) {
    await sendStatusCardToEmployee(
      nctx, employee.teams_id,
      ctx.activity.from.id, ctx.activity.conversation.id,
      updated.type, displayDate, "Rejected", ctx.userName, reason
    );
  }
}