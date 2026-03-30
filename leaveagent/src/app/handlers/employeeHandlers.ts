import {
  buildMyRequestsCard,
  buildLeaveBalanceCard,
  buildMyStatusCard,
  buildSuccessCard,
  buildErrorCard,
  formatDisplayDate,
} from "../cards";
import {
  getLeaveRequestsByEmployee,
  getConversationRef,
  deleteLeaveRequest,
  findEmployee,
  getLeaveBalance,
  getMonthlyPendingRequests,
} from "../postgresManager";
import { canDelete, canEdit } from "../roleGuard";
import { CommandContext } from "./sharedHandlers";

// ── my requests ────────────────────────────────────────────────────────────

export async function handleMyRequests(ctx: CommandContext): Promise<void> {
  const all     = ctx.cmd.includes("all");
  const records = await getLeaveRequestsByEmployee(ctx.userName);
  const shown   = all ? records : records.slice(0, 5);
  await ctx.send(buildMyRequestsCard(ctx.userName, shown as any[]));
}

// ── my balance ─────────────────────────────────────────────────────────────

export async function handleMyBalance(ctx: CommandContext): Promise<void> {
  const employee = await findEmployee(ctx.userName);
  if (!employee) {
    await ctx.send(buildErrorCard(`I couldn't find ${ctx.userName} in the employee directory. Please ask HR to add you.`));
    return;
  }

  const allRecords  = await getLeaveRequestsByEmployee(ctx.userName);
  const pendingDays = allRecords
    .filter((r) => r.status === "Pending" && r.type !== "WFH")
    .reduce((sum, r) => sum + r.days_count, 0);

  const balance = await getLeaveBalance(ctx.userName);
  await ctx.send(buildLeaveBalanceCard(ctx.userName, balance, pendingDays, employee.carry_forward));
}

// ── my status [date] ───────────────────────────────────────────────────────

export async function handleMyStatus(ctx: CommandContext): Promise<void> {
  const dateMatch = ctx.userMessage.match(/(\d{4}-\d{2}-\d{2})/);
  if (!dateMatch) {
    await ctx.send(buildErrorCard("Please provide a date: `my status 2026-04-06`"));
    return;
  }

  const date    = dateMatch[1];
  const records = await getLeaveRequestsByEmployee(ctx.userName);
  const record  = records.find((r) => r.date === date);

  if (!record) {
    await ctx.send(buildErrorCard(`No request found for ${formatDisplayDate(date)}.`));
    return;
  }

  await ctx.send(buildMyStatusCard(record as any));
}

// ── delete request [date] ──────────────────────────────────────────────────

export async function handleDeleteRequest(ctx: CommandContext): Promise<void> {
  const dateMatch = ctx.userMessage.match(/(\d{4}-\d{2}-\d{2})/);
  if (!dateMatch) {
    await ctx.send(buildErrorCard("Please provide a date: `delete request 2026-04-06`"));
    return;
  }

  const date    = dateMatch[1];
  const records = await getLeaveRequestsByEmployee(ctx.userName);
  const record  = records.find((r) => r.date === date);

  if (!record) {
    await ctx.send(buildErrorCard(`No request found for ${formatDisplayDate(date)}.`));
    return;
  }

  if (!canDelete(ctx.userName, record.employee, record.status)) {
    await ctx.send(buildErrorCard(
      record.status !== "Pending"
        ? `This request has already been ${record.status.toLowerCase()} and cannot be deleted. Contact HR if needed.`
        : "You can only delete your own requests."
    ));
    return;
  }

  const deleted = await deleteLeaveRequest(ctx.userName, date, ctx.userName, "Deleted by employee");
  if (deleted) {
    await ctx.send(buildSuccessCard("Request Deleted", `Your ${record.type} request for ${formatDisplayDate(date)} has been deleted.`));
  } else {
    await ctx.send(buildErrorCard("Could not delete the request. Please try again."));
  }
}

// ── edit request [date] ────────────────────────────────────────────────────

export async function handleEditRequest(
  ctx:              CommandContext,
  savePendingRequest: Function,
  getPendingRequest:  Function
): Promise<void> {
  const dateMatch = ctx.userMessage.match(/(\d{4}-\d{2}-\d{2})/);
  if (!dateMatch) {
    await ctx.send(buildErrorCard("Please provide a date: `edit request 2026-04-06`"));
    return;
  }

  const date    = dateMatch[1];
  const records = await getLeaveRequestsByEmployee(ctx.userName);
  const record  = records.find((r) => r.date === date);

  if (!record) {
    await ctx.send(buildErrorCard(`No request found for ${formatDisplayDate(date)}.`));
    return;
  }

  if (!canEdit(ctx.userName, record.employee, record.status)) {
    await ctx.send(buildErrorCard(
      record.status !== "Pending"
        ? `This request has already been ${record.status.toLowerCase()} and cannot be edited.`
        : "You can only edit your own requests."
    ));
    return;
  }

  // Save as pending with edit history so next message is treated as edit
  await savePendingRequest({
    userId:       ctx.userId,
    userName:     ctx.userName,
    intent:       record.type,
    date:         record.date,
    end_date:     record.end_date ?? undefined,
    duration:     record.duration,
    days_count:   record.days_count,
    reason:       record.reason ?? undefined,
    balanceResult: { requested: record.days_count, balance: 999, granted: record.days_count, lop: 0, hasLop: false },
    history: [
      { role: "user",      content: `I want to edit my ${record.type} on ${record.date}` },
      { role: "assistant", content: "What would you like to change?" },
    ],
  });

  await ctx.send("What would you like to change? You can say:\n- \"Change date to 25th April\"\n- \"Make it a half day\"\n- \"Change to sick leave\"\n- \"Add reason: feeling unwell\"");
}
