import dotenv from "dotenv";
dotenv.config({ path: "env/.env.local" });
dotenv.config({ path: ".env" });

import { App }            from "@microsoft/teams.apps";
import { DevtoolsPlugin } from "@microsoft/teams.dev";
import { parseLeaveIntent } from "./app/groqParser";

import {
  findEmployee,
  addLeaveRequest,
  updateLeaveStatus,
  isDuplicateRequest,
  isOverlappingLeave,
  getTodaysAbsences,
  getLeaveRequestsByEmployee,
  saveConversationRef,
  getConversationRef,
  countWorkingDays,
  checkLeaveBalance,
  getHolidays,
  isHoliday,
  getLeaveBalance,
  savePendingRequest,
  getPendingRequest,
  clearPendingRequest,
} from "./app/postgresManager";

import {
  buildApprovalCardContent,
  buildApprovedCardContent,
  buildRejectedCardContent,
  buildConfirmationCard,
  buildStatusCardContent,
  buildAlreadyProcessedCardContent,
  buildRejectionReasonPromptCard,
  buildPreviewCard,
  buildCancelledCard,
  formatDisplayDate,
} from "./app/cards";

import {
  sendApprovalCardToApprover,
  sendStatusCardToEmployee,
  sendApprovalAnnouncement,
  sendWorkforceCardToManager,
  sendHRAlert,
  NotificationContext,
} from "./app/notificationServices";

import { getRoleContext }    from "./app/roleGuard";
import { routeCommand }      from "./app/commandRouter";
import { startSchedulers, runStartupChecks } from "./app/schedulers";
import { CommandContext }    from "./app/handlers/sharedHandlers";

// ── App ────────────────────────────────────────────────────────────────────

const app = new App({
  clientId:     process.env.MICROSOFT_APP_ID      ?? process.env.BOT_ID       ?? "",
  clientSecret: process.env.MICROSOFT_APP_PASSWORD ?? process.env.BOT_PASSWORD ?? "",
  tenantId:     process.env.MICROSOFT_APP_TENANT_ID ?? "",
  plugins:      [new DevtoolsPlugin()],
});

// ── Message Handler ────────────────────────────────────────────────────────

app.on("message", async ({ activity, send, api }) => {
  const userMessage = (activity.text ?? "").replace(/<[^>]+>/g, "").trim();
  const userId      = activity.from.id;
  const userName    = activity.from.name ?? "Employee";

  console.log(`[LeaveAgent] "${userMessage}" from ${userName} (${userId})`);

  await saveConversationRef({
    userId,
    userName,
    conversationId: activity.conversation.id,
    serviceUrl:     activity.serviceUrl,
    tenantId:       activity.conversation.tenantId,
    botId:          activity.recipient?.id ?? "bot",
  });

  const nctx: NotificationContext = { api, botId: activity.recipient?.id ?? "" };

  // ── Intercept card button clicks misrouted as messages ───────────────────
  const activityValue = (activity as any).value;
  if (activityValue?.action) {
    console.log("[FLOW] CARD ACTION FLOW triggered");
    const normalizedData = activityValue?.action?.data ?? activityValue;
    const handled = await handleCardAction(normalizedData, userName, userId, activity, send, api, nctx);
    console.log("[FLOW] handleCardAction returned:", handled);
    if (handled) {
          console.log("[FLOW] Card handled → STOP");
      return;
    }
    console.log("[FLOW] Card NOT handled → FALLBACK to text flow ❌");

  }

  // ── Role + Command Router ─────────────────────────────────────────────────
  const role = await getRoleContext(userName);
  const cmd  = userMessage.toLowerCase();
  const ctx: CommandContext = { activity, send, api, userName, userId, userMessage, cmd, role };

  // ── Edit mode (pending with history) ─────────────────────────────────────
  const existingPending = await getPendingRequest(userId);
  if (existingPending?.history?.length) {
    await handleEditMode(ctx, existingPending, nctx);
    return;
  }

  // ── Command table ─────────────────────────────────────────────────────────
  const routed = await routeCommand(ctx, nctx, { savePendingRequest, getPendingRequest });
  if (routed) return;

  // ── AI intent parsing ─────────────────────────────────────────────────────
  await handleLeaveRequest(ctx, nctx);
});

// ── Card Action Handler ────────────────────────────────────────────────────

app.on("card.action", async ({ activity, send, api }) => {
  const data            = (activity.value as any)?.action?.data ?? (activity.value as any);
  const approverName    = activity.from.name ?? "Approver";
  const approverTeamsId = activity.from.id;
  const nctx: NotificationContext = { api, botId: activity.recipient?.id ?? "" };

  console.log(`[LeaveAgent] card.action from ${approverName}:`, data);

  if (!data?.action) {
    return { statusCode: 200, type: "application/vnd.microsoft.card.adaptive", value: buildAlreadyProcessedCardContent() as any } as any;
  }

  // Preview actions
  if (data.action === "preview_confirm") {
    await submitRequest(activity.from.id, approverName, activity, send, api, nctx);
    return { statusCode: 200, type: "application/vnd.microsoft.card.adaptive", value: buildAlreadyProcessedCardContent() as any } as any;
  }

  if (data.action === "preview_edit") {
    const pending = await getPendingRequest(activity.from.id);
    if (!pending) { await send("No pending request. Please start a new one."); return { statusCode: 200 } as any; }
    pending.history = [
      { role: "user",      content: `I want to request ${pending.intent} on ${pending.date}` },
      { role: "assistant", content: "What would you like to change?" },
    ];
    await savePendingRequest(pending);
    await send("What would you like to change?");
    return { statusCode: 200, type: "application/vnd.microsoft.card.adaptive", value: buildAlreadyProcessedCardContent() as any } as any;
  }

  if (data.action === "preview_cancel") {
    await clearPendingRequest(activity.from.id);
    await send(buildCancelledCard());
    return { statusCode: 200, type: "application/vnd.microsoft.card.adaptive", value: buildAlreadyProcessedCardContent() as any } as any;
  }

  // Approver actions
  if (!data?.employeeName || !data?.date) {
    return { statusCode: 200, type: "application/vnd.microsoft.card.adaptive", value: buildAlreadyProcessedCardContent() as any } as any;
  }

  const { action, employeeName, date, requestType = "WFH" } = data;
  const displayDate = formatDisplayDate(date);

  // Reject — show reason prompt card
  if (action === "reject") {
    return {
      statusCode: 200,
      type:       "application/vnd.microsoft.card.adaptive",
      value:      buildRejectionReasonPromptCard(employeeName, date, requestType, displayDate) as any,
    } as any;
  }

  // Confirm reject with reason
  if (action === "confirm_reject") {
    const reason  = data.rejectionReason || "No reason provided";
    const updated = await updateLeaveStatus(employeeName, date, "Rejected", approverName, reason);
    if (!updated) return { statusCode: 200, type: "application/vnd.microsoft.card.adaptive", value: buildAlreadyProcessedCardContent() as any } as any;

    const employee = await findEmployee(employeeName);
    if (employee?.teams_id) {
      await sendStatusCardToEmployee(nctx, employee.teams_id, approverTeamsId, activity.conversation.id, requestType, displayDate, "Rejected", approverName, reason, send);
    }
    return {
      statusCode: 200, type: "application/vnd.microsoft.card.adaptive",
      value: buildRejectedCardContent({ employeeName, requestType, date, displayDate }, approverName, reason) as any,
    } as any;
  }

  // Approve
  if (action === "approve") {
    const updated = await updateLeaveStatus(employeeName, date, "Approved", approverName);
    if (!updated) return { statusCode: 200, type: "application/vnd.microsoft.card.adaptive", value: buildAlreadyProcessedCardContent() as any } as any;

    const employee    = await findEmployee(employeeName);
    const allRecords  = await getLeaveRequestsByEmployee(employeeName);
    const leaveRecord = allRecords.find((r) => r.date === date);
    const dur         = leaveRecord?.duration  ?? "full_day";
    const days        = leaveRecord?.days_count ?? 1;

    if (employee?.teams_id) {
      await sendStatusCardToEmployee(nctx, employee.teams_id, approverTeamsId, activity.conversation.id, requestType, displayDate, "Approved", approverName, undefined, send);
    }
    await sendApprovalAnnouncement(nctx, employeeName, requestType, date, displayDate, leaveRecord?.end_date);
    if (employee) await sendWorkforceCardToManager(nctx, employee, employeeName, requestType, date, leaveRecord, dur, days, approverName, approverTeamsId);
    await sendHRAlert(nctx, "approved", employeeName, requestType, displayDate, approverName);

    return {
      statusCode: 200, type: "application/vnd.microsoft.card.adaptive",
      value: buildApprovedCardContent({ employeeName, requestType, date, displayDate }, approverName) as any,
    } as any;
  }

  return { statusCode: 200 } as any;
});

// ── Welcome ────────────────────────────────────────────────────────────────

app.on("install.add", async ({ send }) => {
  await send(
    "Hi! I'm LeaveAgent 👋\n\n" +
    "Submit leave requests naturally:\n" +
    "• 'WFH tomorrow'\n• 'Sick today'\n• 'Leave from 20th to 25th April'\n\n" +
    "Type `help` to see all available commands."
  );
});

// ── Internal Helpers ───────────────────────────────────────────────────────

async function handleCardAction(
  data: any, userName: string, userId: string,
  activity: any, send: Function, api: any, nctx: NotificationContext
): Promise<boolean> {


  console.log("------ CARD ACTION START ------");
  console.log("[CARD] Full payload:", JSON.stringify(data));
  console.log("[CARD] action:", data.action);

  const { action, employeeName, date, requestType = "WFH" } = data;


  if (action === "preview_confirm") {
  await submitRequest(userId, userName, activity, send, api, nctx);
  return true;
  }
if (action === "preview_edit") {
  const pending = await getPendingRequest(userId);
  if (!pending) { await send("No pending request. Please start a new one."); return true; }
  pending.history = [
    { role: "user",      content: `I want to request ${pending.intent} on ${pending.date}` },
    { role: "assistant", content: "What would you like to change?" },
  ];
  await savePendingRequest(pending);
  await send("What would you like to change?");
  return true;
}
if (action === "preview_cancel") {
  await clearPendingRequest(userId);
  await send(buildCancelledCard());
  return true;
}



  if (action !== "approve" && action !== "reject" && action !== "confirm_reject") return false;

if (!employeeName || !date) {
  await send("Invalid request data.");
  return true;
}

  const displayDate = formatDisplayDate(date);

  if (action === "reject") {
    await send({
      type: "message",
      attachments: [{ contentType: "application/vnd.microsoft.card.adaptive", content: buildRejectionReasonPromptCard(employeeName, date, requestType, displayDate) }],
    } as any);
    return true;
  }

  if (action === "confirm_reject") {
    const reason  = data.rejectionReason || "No reason provided";
    const updated = await updateLeaveStatus(employeeName, date, "Rejected", userName, reason);
    if (!updated) { await send("This request has already been processed."); return true; }
    await send(`Request rejected for ${employeeName} on ${displayDate}. Reason: ${reason}`);
    const employee = await findEmployee(employeeName);
    if (employee?.teams_id) {
      await sendStatusCardToEmployee(nctx, employee.teams_id, userId, activity.conversation.id, requestType, displayDate, "Rejected", userName, reason, send);
    }
    return true;
  }

  // approve
  const updated = await updateLeaveStatus(employeeName, date, "Approved", userName);
  if (!updated) { await send("This request has already been processed."); return true; }

  await send(`Request Approved for ${employeeName} on ${displayDate}.`);

  const employee    = await findEmployee(employeeName);
  const allRecords  = await getLeaveRequestsByEmployee(employeeName);
  const leaveRecord = allRecords.find((r) => r.date === date);
  const dur         = leaveRecord?.duration  ?? "full_day";
  const days        = leaveRecord?.days_count ?? 1;

  if (employee?.teams_id) {
    await sendStatusCardToEmployee(nctx, employee.teams_id, userId, activity.conversation.id, requestType, displayDate, "Approved", userName, undefined, send);
    await sendApprovalAnnouncement(nctx, employeeName, requestType, date, displayDate, leaveRecord?.end_date);
    await sendWorkforceCardToManager(nctx, employee, employeeName, requestType, date, leaveRecord, dur, days, userName, userId);
    await sendHRAlert(nctx, "approved", employeeName, requestType, displayDate, userName);
  }

  return true;
}

async function handleEditMode(ctx: CommandContext, existingPending: any, nctx: NotificationContext): Promise<void> {
  await ctx.send("Processing your updated request...");

  const intent = await parseLeaveIntent(ctx.userMessage, existingPending.history);
  if (intent.needs_clarification || intent.intent === "UNKNOWN") {
    existingPending.history.push(
      { role: "user",      content: ctx.userMessage },
      { role: "assistant", content: intent.clarification_question ?? "" }
    );
    await savePendingRequest(existingPending);
    await ctx.send(intent.clarification_question ?? "Could you clarify?");
    return;
  }

  const employee = await findEmployee(ctx.userName);
  if (!employee) { await ctx.send("Employee not found. Please ask HR to add you."); return; }

  const displayDate    = formatDisplayDate(intent.date);
  const displayEndDate = intent.end_date ? formatDisplayDate(intent.end_date) : null;
  const durationLabel  = intent.duration === "half_day" ? "Half Day" : intent.duration === "multi_day" ? "Multiple Days" : "Full Day";
  const daysCount = await countWorkingDays(intent.date, intent.end_date);
  const balanceResult  = await checkLeaveBalance(employee, daysCount, intent.intent, intent.date, intent.end_date);

  await savePendingRequest({
    userId:       ctx.userId,
    userName:     ctx.userName,
    intent:       intent.intent,
    date:         intent.date,
    end_date:     intent.end_date,
    duration:     intent.duration,
    days_count:   daysCount,
    reason:       intent.reason,
    balanceResult,
    history:      [],
  });

  await ctx.send(buildPreviewCard({
    employeeName: ctx.userName,
    requestType:  intent.intent,
    date:         intent.date,
    displayDate,
    endDate:      displayEndDate,
    daysCount,
    duration:     durationLabel,
    reason:       intent.reason,
    balanceResult,
  }));
}

// Clear any stale pending before saving new one
/*await clearPendingRequest(ctx: CommandContext);
 
await savePendingRequest({
  userId:       ctx.userId,
  ...
});*/
async function handleLeaveRequest(ctx: CommandContext, nctx: NotificationContext): Promise<void> {
  const onBehalfMatch = ctx.userMessage.match(/(?:leave|wfh|sick).*\bfor\s+([A-Z][a-z]+(?:\s+[A-Z][a-z]+)*)/);
  if (onBehalfMatch && onBehalfMatch[1].toLowerCase() !== ctx.userName.toLowerCase() && !ctx.role.isHR) {
    await ctx.send(`You can only submit requests for yourself. Please ask ${onBehalfMatch[1]} to submit their own request.`);
    return;
  }

  await ctx.send("Processing your request...");

  const intent = await parseLeaveIntent(ctx.userMessage);
  console.log(`[LeaveAgent] Intent:`, JSON.stringify(intent));

  if (intent.needs_clarification || intent.intent === "UNKNOWN") {
    await ctx.send(intent.clarification_question ?? "Could you rephrase? Try: 'WFH tomorrow', 'Sick today', or 'Leave from 20th to 25th'.");
    return;
  }

  if (intent.is_third_party && !ctx.role.isHR) {
    await ctx.send(`You can only submit requests for yourself.`);
    return;
  }

  if (await isDuplicateRequest(ctx.userName, intent.date)) {
    await ctx.send(`You already have a request for ${formatDisplayDate(intent.date)}. Type 'my requests' to view it.`);
    return;
  }

  const overlap = await isOverlappingLeave(ctx.userName, intent.date, intent.end_date ?? intent.date);
  if (overlap.overlaps) {
    await ctx.send(`Your request overlaps with an existing leave (${overlap.conflictDate}).`);
    return;
  }

  if (!intent.end_date || intent.date === intent.end_date) {
    const holiday = await isHoliday(intent.date);
    if (holiday) {
      const holidays = await getHolidays();
      const h = holidays.find((hol) => hol.date === intent.date);
      await ctx.send(`🎉 ${formatDisplayDate(intent.date)} is already a public holiday${h ? ` — ${h.name}` : ""}! No leave request needed.`);
      return;
    }
  }

  const employee = await findEmployee(ctx.userName);
  if (!employee) { await ctx.send(`I couldn't find ${ctx.userName} in the directory. Please ask HR to add you.`); return; }

  const displayDate    = formatDisplayDate(intent.date);
  const displayEndDate = intent.end_date ? formatDisplayDate(intent.end_date) : null;
  const durationLabel  = intent.duration === "half_day" ? "Half Day" : intent.duration === "multi_day" ? "Multiple Days" : "Full Day";
  const daysCount = await countWorkingDays(intent.date, intent.end_date);
  const balanceResult  = await checkLeaveBalance(employee, daysCount, intent.intent, intent.date, intent.end_date);

  if ((balanceResult as any).needsCarryForward) {
    await ctx.send("Your carry forward hasn't been calculated yet. Please request January leaves after December 25th.");
    return;
  }

  if (balanceResult.hasLop) {
    await ctx.send(
      `Balance: ${balanceResult.balance.toFixed(1)} day(s)\n` +
      `Requested: ${balanceResult.requested} day(s)\n` +
      `Granted: ${balanceResult.granted} day(s) | LOP: ${balanceResult.lop} day(s) — contact HR.`
    );
  }

  await savePendingRequest({
    userId:       ctx.userId,
    userName:     ctx.userName,
    intent:       intent.intent,
    date:         intent.date,
    end_date:     intent.end_date,
    duration:     intent.duration,
    days_count:   daysCount,
    reason:       intent.reason,
    balanceResult,
    history:      [],
  });

  await ctx.send(buildPreviewCard({
    employeeName: ctx.userName,
    requestType:  intent.intent,
    date:         intent.date,
    displayDate,
    endDate:      displayEndDate,
    daysCount,
    duration:     durationLabel,
    reason:       intent.reason,
    balanceResult,
  }));
}

async function submitRequest(
  userId:   string,
  userName: string,
  activity: any,
  send:     Function,
  api:      any,
  nctx:     NotificationContext
): Promise<void> {
  const pending = await getPendingRequest(userId);
  if (!pending) { await send("No pending request. Please start a new one."); return; }

  const employee = await findEmployee(userName);
  if (!employee) { await send(`Employee not found. Please ask HR to add you.`); await clearPendingRequest(userId); return; }

  const { intent, date, end_date, duration, days_count, reason, balanceResult } = pending;
  const displayDate    = formatDisplayDate(date);
  const displayEndDate = end_date ? formatDisplayDate(end_date) : null;
  const durationLabel  = duration === "half_day" ? "Half Day" : duration === "multi_day" ? "Multiple Days" : "Full Day";
  const isTeamLead     = employee.role === "teamlead";
  const approverName   = isTeamLead ? employee.manager          : employee.teamlead;
  const approverTeamsId = isTeamLead ? employee.manager_teams_id : employee.teamlead_teams_id;

  await addLeaveRequest({
    employee:   userName,
    email:      employee.email,
    type:       intent,
    date,
    end_date:   end_date ?? undefined,
    duration,
    days_count,
    reason:     reason ?? undefined,
    status:     "Pending",
  });

  await clearPendingRequest(userId);

  await send(buildConfirmationCard(userName, intent, displayDate, durationLabel, displayEndDate, days_count, reason, balanceResult));

  await sendHRAlert(nctx, "submitted", userName, intent, displayDate, userName);

  const approvalCardPayload = buildApprovalCardContent({
    employeeName:  userName,
    employeeEmail: employee.email,
    requestType:   intent,
    date,
    displayDate,
    duration:      durationLabel,
    endDate:       displayEndDate,
    daysCount:     days_count,
    reason,
    balanceResult,
  });

  await sendApprovalCardToApprover(nctx, approverTeamsId ?? "", approverName ?? "Approver", approvalCardPayload, send);
}

// ── Start ──────────────────────────────────────────────────────────────────

(async () => {
  await app.start(+(process.env.PORT ?? 3978));
  console.log(`\nLeaveAgent running on port ${process.env.PORT ?? 3978}\n`);

  runStartupChecks();

  startSchedulers({
    sendApproverReminder: async (month, year) => {
      const { getMonthlyPendingRequests, getAllEmployees } = await import("./app/postgresManager.js");
      const { sendApproverReminders }                     = await import("./app/notificationServices.js");
      const records = await getMonthlyPendingRequests();
      const allEmps = await getAllEmployees();
      const label   = new Date(year, month - 1, 1).toLocaleDateString("en-IN", { month: "long", year: "numeric" });
      const groups: Record<string, any> = {};
      for (const r of records) {
        const emp = allEmps.find((e) => e.name.toLowerCase() === r.employee.toLowerCase());
        if (!emp) continue;
        const tid  = emp.role === "teamlead" ? emp.manager_teams_id  : emp.teamlead_teams_id;
        const name = emp.role === "teamlead" ? emp.manager           : emp.teamlead;
        if (!tid || !name) continue;
        if (!groups[tid]) groups[tid] = { approverName: name, approverTeamsId: tid, records: [] };
        groups[tid].records.push(r);
      }
      await sendApproverReminders(Object.values(groups), label);
    },
    sendHRTakeover: async (month, year) => {
      const { getMonthlyPendingRequests } = await import("./app/postgresManager.js");
      const { sendHRTakeover }            = await import("./app/notificationServices.js");
      const records = await getMonthlyPendingRequests();
      const label   = new Date(year, month - 1, 1).toLocaleDateString("en-IN", { month: "long", year: "numeric" });
      await sendHRTakeover(records as any[], label);
    },
  });
})();
