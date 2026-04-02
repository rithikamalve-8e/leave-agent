import dotenv from "dotenv";
dotenv.config({ path: "env/.env.local" });
dotenv.config({ path: ".env" });

import { App }             from "@microsoft/teams.apps";
import { DevtoolsPlugin }  from "@microsoft/teams.dev";
import { parseLeaveIntent } from "./app/groqParser";

// ── DB imports — dead ones (getTodaysAbsences, getConversationRef, getLeaveBalance) removed ──
import {
  findEmployee,
  addLeaveRequest,
  updateLeaveStatus,
  isDuplicateRequest,
  isOverlappingLeave,
  getLeaveRequestsByEmployee,
  saveConversationRef,
  countWorkingDays,
  checkLeaveBalance,
  getHolidays,
  isHoliday,
  savePendingRequest,
  getPendingRequest,
  getLeaveRequestStatus,
  clearPendingRequest,
} from "./app/postgresManager";

// ── Card imports — dead one (buildStatusCardContent) removed ──
import {
  buildApprovalCardContent,
  buildApprovedCardContent,
  buildRejectedCardContent,
  buildConfirmationCard,
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

import { getRoleContext }   from "./app/roleGuard";
import { routeCommand }     from "./app/commandRouter";
import { startSchedulers, runStartupChecks } from "./app/schedulers";
import { CommandContext }   from "./app/handlers/sharedHandlers";

// ── App ────────────────────────────────────────────────────────────────────

const app = new App({
  clientId:     process.env.MICROSOFT_APP_ID      ?? process.env.BOT_ID       ?? "",
  clientSecret: process.env.MICROSOFT_APP_PASSWORD ?? process.env.BOT_PASSWORD ?? "",
  tenantId:     process.env.MICROSOFT_APP_TENANT_ID ?? "",
  plugins:      [new DevtoolsPlugin()],
});

// ── In-flight reject guard ─────────────────────────────────────────────────
// Prevents double-clicking Reject from sending two reason-prompt cards.
// Keyed as `${employeeName}:${date}`. Cleared after 10 min in case the
// approver abandons without submitting or cancelling.
const pendingRejections = new Set<string>();

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

  console.log(`[card.action] from ${approverName}:`, data);

  if (!data?.action) {
    return { statusCode: 200, type: "application/vnd.microsoft.card.adaptive", value: buildAlreadyProcessedCardContent() } as any;
  }

  // ── Preview actions (employee-side) ─────────────────────────────────────

  if (data.action === "preview_confirm") {
    await submitRequest(activity.from.id, approverName, activity, send, api, nctx);
    return { statusCode: 200, type: "application/vnd.microsoft.card.adaptive", value: buildAlreadyProcessedCardContent() } as any;
  }

  if (data.action === "preview_edit") {
    const pending = await getPendingRequest(activity.from.id);
    if (!pending) {
      await send("No pending request. Please start a new one.");
      return { statusCode: 200 } as any;
    }
    pending.history = [
      { role: "user",      content: `I want to request ${pending.intent} on ${pending.date}` },
      { role: "assistant", content: "What would you like to change?" },
    ];
    await savePendingRequest(pending);
    await send("What would you like to change? (e.g. 'Change date to 5th April', 'Make it half day')");
    // card.action path: return the static card inline — Teams replaces the card automatically
    return { statusCode: 200, type: "application/vnd.microsoft.card.adaptive", value: buildAlreadyProcessedCardContent() } as any;
  }

  if (data.action === "preview_cancel") {
    await clearPendingRequest(activity.from.id);
    await send(buildCancelledCard());
    return { statusCode: 200, type: "application/vnd.microsoft.card.adaptive", value: buildAlreadyProcessedCardContent() } as any;
  }

  // ── Approver actions ─────────────────────────────────────────────────────

  if (!data?.employeeName || !data?.date) {
    return { statusCode: 200, type: "application/vnd.microsoft.card.adaptive", value: buildAlreadyProcessedCardContent() } as any;
  }

  const { action, employeeName, date, requestType = "WFH" } = data;
  const displayDate = formatDisplayDate(date);

  // REJECT — idempotency guard: prevent double reason-prompt cards
  if (action === "reject") {
    const key = `${employeeName}:${date}`;
    if (pendingRejections.has(key)) {
      console.log(`[card.action] reject already in-flight for ${key} — ignoring`);
      return { statusCode: 200, type: "application/vnd.microsoft.card.adaptive", value: buildAlreadyProcessedCardContent() } as any;
    }
    pendingRejections.add(key);
    setTimeout(() => pendingRejections.delete(key), 10 * 60 * 1000);

    // FIX: unwrap .attachments[0].content — card.action needs raw IAdaptiveCard, not CardActivity wrapper
    return {
      statusCode: 200,
      type:       "application/vnd.microsoft.card.adaptive",
      value:      buildRejectionReasonPromptCard(employeeName, date, requestType, displayDate).attachments[0].content,
    } as any;
  }

  // CANCEL REJECT — clear guard, restore approval card inline
  if (action === "cancel_reject") {
    const key = `${employeeName}:${date}`;
    pendingRejections.delete(key);

    const allRecords  = await getLeaveRequestsByEmployee(employeeName);
    const leaveRecord = allRecords.find((r: any) => r.date === date);
    if (!leaveRecord) {
      return { statusCode: 200, type: "application/vnd.microsoft.card.adaptive", value: buildAlreadyProcessedCardContent() } as any;
    }
    const employee = await findEmployee(employeeName);
    // Return the approval card inline so the approver can still approve/reject
    return {
      statusCode: 200,
      type:       "application/vnd.microsoft.card.adaptive",
      value:      buildApprovalCardContent({
        employeeName,
        employeeEmail: employee?.email ?? "",
        requestType,
        date,
        displayDate,
        duration:  leaveRecord.duration === "half_day" ? "Half Day" : leaveRecord.duration === "multi_day" ? "Multiple Days" : "Full Day",
        daysCount: leaveRecord.days_count,
        reason:    leaveRecord.reason,
      }),
    } as any;
  }

  // CONFIRM REJECT
  if (action === "confirm_reject") {
    // Clear guard — action is now finalised
    pendingRejections.delete(`${employeeName}:${date}`);

    const existing = await getLeaveRequestStatus(employeeName, date);
    if (existing && existing.status !== "Pending") {
      return { statusCode: 200, type: "application/vnd.microsoft.card.adaptive", value: buildAlreadyProcessedCardContent() } as any;
    }

    const reason  = data.rejectionReason || "No reason provided";
    const updated = await updateLeaveStatus(employeeName, date, "Rejected", approverName, reason);
    if (!updated) {
      return { statusCode: 200, type: "application/vnd.microsoft.card.adaptive", value: buildAlreadyProcessedCardContent() } as any;
    }

    const employee = await findEmployee(employeeName);
    if (employee?.teams_id) {
      await sendStatusCardToEmployee(nctx, employee.teams_id, approverTeamsId, activity.conversation.id, requestType, displayDate, "Rejected", approverName, reason, send);
    } else {
      // FIX: employee has no teams_id — fall back to logging so rejection isn't silent
      console.warn(`[card.action] confirm_reject: no teams_id for ${employeeName} — employee not notified`);
    }

    return {
      statusCode: 200,
      type:       "application/vnd.microsoft.card.adaptive",
      value:      buildRejectedCardContent({ employeeName, requestType, date, displayDate }, approverName, reason),
    } as any;
  }

  // APPROVE
  if (action === "approve") {
    const existing = await getLeaveRequestStatus(employeeName, date);
    if (existing && existing.status !== "Pending") {
      return { statusCode: 200, type: "application/vnd.microsoft.card.adaptive", value: buildAlreadyProcessedCardContent() } as any;
    }

    const updated = await updateLeaveStatus(employeeName, date, "Approved", approverName);
    if (!updated) {
      return { statusCode: 200, type: "application/vnd.microsoft.card.adaptive", value: buildAlreadyProcessedCardContent() } as any;
    }

    const employee    = await findEmployee(employeeName);
    const allRecords  = await getLeaveRequestsByEmployee(employeeName);
    const leaveRecord = allRecords.find((r: any) => r.date === date);
    const dur         = leaveRecord?.duration  ?? "full_day";
    const days        = leaveRecord?.days_count ?? 1;

    if (employee?.teams_id) {
      await sendStatusCardToEmployee(nctx, employee.teams_id, approverTeamsId, activity.conversation.id, requestType, displayDate, "Approved", approverName, undefined, send);
    } else {
      console.warn(`[card.action] approve: no teams_id for ${employeeName} — employee not notified`);
    }

    await sendApprovalAnnouncement(nctx, employeeName, requestType, date, displayDate, leaveRecord?.end_date,leaveRecord?.duration);
    if (employee) await sendWorkforceCardToManager(nctx, employee, employeeName, requestType, date, leaveRecord, dur, days, approverName, approverTeamsId);
    await sendHRAlert(nctx, "approved", employeeName, requestType, displayDate, approverName);

    return {
      statusCode: 200,
      type:       "application/vnd.microsoft.card.adaptive",
      value:      buildApprovedCardContent({ employeeName, requestType, date, displayDate }, approverName),
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

// ── handleCardAction — message-path fallback ───────────────────────────────
// Teams sometimes misroutes card button clicks through the `message` handler
// instead of `card.action`. This function handles those cases.
// It mirrors card.action logic but uses api.http.put for card replacement
// instead of inline return values (which only work in card.action).

async function handleCardAction(
  data: any, userName: string, userId: string,
  activity: any, send: Function, api: any, nctx: NotificationContext
): Promise<boolean> {

  console.log("------ CARD ACTION START (message path) ------");
  console.log("[CARD] payload:", JSON.stringify(data));
  console.log("[CARD] action:", data.action);
  // Diagnostic: tells us which ID replaceActivityCard will target
  console.log("[CARD] replyToId:", activity.replyToId, "| id:", activity.id);

  const { action, employeeName, date, requestType = "WFH" } = data;

  // ── Preview actions ───────────────────────────────────────────────────────

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

    // FIX: use stored previewCardActivityId for reliable targeting.
    // If not stored, fall back to replyToId then activity.id.
    const targetCardId = pending.previewCardActivityId ?? activity.replyToId ?? activity.id;
    await replaceActivityCard(api, activity, {
      $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
      type: "AdaptiveCard", version: "1.4",
      body: [
        { type: "TextBlock", text: "Editing Request", weight: "Bolder", size: "Large", color: "Accent" },
        { type: "TextBlock", text: "Type your changes below (e.g. 'Change date to 5th April', 'Make it half day').", wrap: true, size: "Small", isSubtle: true },
      ],
    }, targetCardId);

    await send("What would you like to change? (e.g. 'Change date to 5th April', 'Make it half day')");
    return true;
  }

  if (action === "preview_cancel") {
    await clearPendingRequest(userId);
    await send(buildCancelledCard());
    return true;
  }

  // ── Approver actions ─────────────────────────────────────────────────────

  if (
    action !== "approve" &&
    action !== "reject" &&
    action !== "confirm_reject" &&
    action !== "cancel_reject"
  ) return false;

  if (!employeeName || !date) {
    await send("Invalid request data.");
    return true;
  }

  const displayDate = formatDisplayDate(date);

  // Idempotency guard — approve and confirm_reject write to DB
  if (action === "approve" || action === "confirm_reject") {
    const existing = await getLeaveRequestStatus(employeeName, date);
    if (existing && existing.status !== "Pending") {
      console.log(`[CARD] Already processed (${existing.status}) — ignoring`);
      await send(`This request has already been ${existing.status.toLowerCase()}.`);
      return true;
    }
  }

  // REJECT — idempotency guard for reason-prompt
  if (action === "reject") {
    const key = `${employeeName}:${date}`;
    if (pendingRejections.has(key)) {
      console.log(`[CARD] reject already in-flight for ${key} — ignoring duplicate`);
      return true;
    }
    pendingRejections.add(key);
    setTimeout(() => pendingRejections.delete(key), 10 * 60 * 1000);

    // Replace the approve/reject card immediately — buttons gone
    await replaceActivityCard(api, activity, buildAlreadyProcessedCardContent());

    // Send reason-prompt as new message — FIX: unwrap .attachments[0].content
    await send({
      type: "message",
      attachments: [{
        contentType: "application/vnd.microsoft.card.adaptive",
        content: buildRejectionReasonPromptCard(employeeName, date, requestType, displayDate).attachments[0].content,
      }],
    } as any);
    return true;
  }

  // CANCEL REJECT — clear guard, inform approver
  if (action === "cancel_reject") {
    pendingRejections.delete(`${employeeName}:${date}`);
    // In the message path we cannot restore the original card reliably
    // (it was replaced with "already processed"). Inform the approver instead.
    await send(
      `Rejection cancelled for ${employeeName} on ${displayDate}. ` +
      `The request is still pending — use the 'approve leave ${employeeName} ${date}' command to approve, ` +
      `or 'reject leave ${employeeName} ${date} [reason]' to reject via command.`
    );
    return true;
  }

  // CONFIRM REJECT
  if (action === "confirm_reject") {
    pendingRejections.delete(`${employeeName}:${date}`);

    const reason  = data.rejectionReason || "No reason provided";
    const updated = await updateLeaveStatus(employeeName, date, "Rejected", userName, reason);

    if (!updated) {
      console.log("[CARD] confirm_reject: already processed");
      await send("This request has already been processed.");
      return true;
    }

    console.log(`[CARD] Rejected: ${employeeName} on ${date}. Reason: ${reason}`);

    const employee = await findEmployee(employeeName);
    if (employee?.teams_id) {
      await sendStatusCardToEmployee(
        nctx, employee.teams_id, userId,
        activity.conversation.id, requestType, displayDate,
        "Rejected", userName, reason, send
      );
    } else {
      console.warn(`[CARD] confirm_reject: no teams_id for ${employeeName} — employee not notified`);
    }

    // Replace the reason-prompt card with the final rejected summary
    // replyToId here points to the reason-prompt card message — correct target
    await replaceActivityCard(
      api, activity,
      buildRejectedCardContent({ employeeName, requestType, date, displayDate }, userName, reason)
    );
    return true;
  }

  // APPROVE
  if (action === "approve") {
    const updated = await updateLeaveStatus(employeeName, date, "Approved", userName);

    if (!updated) {
      console.log("[CARD] approve: already processed");
      await send("This request has already been processed.");
      return true;
    }

    console.log(`[CARD] Approved: ${employeeName} on ${date}`);

    const employee    = await findEmployee(employeeName);
    const allRecords  = await getLeaveRequestsByEmployee(employeeName);
    const leaveRecord = allRecords.find((r: any) => r.date === date);
    const dur         = leaveRecord?.duration  ?? "full_day";
    const days        = leaveRecord?.days_count ?? 1;

    if (employee?.teams_id) {
      await sendStatusCardToEmployee(
        nctx, employee.teams_id, userId,
        activity.conversation.id, requestType, displayDate,
        "Approved", userName, undefined, send
      );
    } else {
      console.warn(`[CARD] approve: no teams_id for ${employeeName} — employee not notified`);
    }

    await sendApprovalAnnouncement(nctx, employeeName, requestType, date, displayDate, leaveRecord?.end_date);
    if (employee) {
      await sendWorkforceCardToManager(nctx, employee, employeeName, requestType, date, leaveRecord, dur, days, userName, userId);
    }
    await sendHRAlert(nctx, "approved", employeeName, requestType, displayDate, userName);

    await replaceActivityCard(
      api, activity,
      buildApprovedCardContent({ employeeName, requestType, date, displayDate }, userName)
    );
    return true;
  }

  return false;
}

// ── Card Replacement via Bot Framework REST ────────────────────────────────
// api.http is the only real method on the api object (api keys: constructor, http).
// We call the Bot Connector PUT endpoint directly to replace a card in-place.
//
// FIX: serviceUrl trailing-slash normalised to prevent broken URLs.
// FIX: accepts optional explicitCardId — used when caller has the stored
//      previewCardActivityId, bypassing the unreliable activity.replyToId.

async function replaceActivityCard(
  api:             any,
  activity:        any,
  cardContent:     any,
  explicitCardId?: string        // pass pending.previewCardActivityId when available
): Promise<void> {
  // FIX: normalise trailing slash — prevents "smba.trafficmanager.netv3/..." broken URL
  const serviceUrl     = (activity.serviceUrl as string).replace(/\/?$/, "/");
  const conversationId = activity.conversation.id;

  // Priority: explicit stored ID > replyToId > activity.id (last resort)
  const cardActivityId = explicitCardId ?? activity.replyToId ?? activity.id;

  console.log(`[replaceActivityCard] targeting activityId: ${cardActivityId}`);

  if (!cardActivityId || !conversationId) {
    console.warn("[replaceActivityCard] missing cardActivityId or conversationId — skipping");
    return;
  }

  try {
    await api.http.put(
      `${serviceUrl}v3/conversations/${conversationId}/activities/${cardActivityId}`,
      {
        type: "message",
        id:   cardActivityId,
        attachments: [{
          contentType: "application/vnd.microsoft.card.adaptive",
          content:     cardContent,
        }],
      }
    );
    console.log("[replaceActivityCard] success");
  } catch (e: any) {
    // Non-fatal — DB idempotency guard is the real safety net
    console.warn("[replaceActivityCard] failed:", e?.message ?? e);
  }
}

// ── Edit Mode ─────────────────────────────────────────────────────────────

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
  const durationLabel =
  intent.duration === "morning"
    ? "Morning"
    : intent.duration === "afternoon"
    ? "Afternoon"
    : intent.duration === "multi_day"
    ? "Multiple Days"
    : "Full Day";  
  const rawDays = await countWorkingDays(intent.date, intent.end_date);
  const daysCount =
  intent.duration === "morning" || intent.duration === "afternoon"
    ? 0.5
    : rawDays;  
  const balanceResult  = await checkLeaveBalance(employee, daysCount, intent.intent, intent.date, intent.end_date);



  
  // Send updated preview card and capture the activity ID Teams returns
  const sentActivity = await ctx.send(buildPreviewCard({
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

  // FIX: save with history cleared (exits edit mode) AND store new card's
  // activity ID so a subsequent Edit click can reliably replace this card.
  await savePendingRequest({
    userId:               ctx.userId,
    userName:             ctx.userName,
    intent:               intent.intent,
    date:                 intent.date,
    end_date:             intent.end_date,
    duration:             intent.duration,
    days_count:           daysCount,
    reason:               intent.reason,
    balanceResult,
    history:              [],                          // ← cleared: exits edit mode
    previewCardActivityId: sentActivity?.id ?? null,  // ← stored: reliable card targeting
  });
}

// ── Leave Request Handler ─────────────────────────────────────────────────

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

  if (await isDuplicateRequest(ctx.userName, intent.date, intent.duration)) {
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
  const durationLabel =
  intent.duration === "morning"
    ? "Morning"
    : intent.duration === "afternoon"
    ? "Afternoon"
    : intent.duration === "multi_day"
    ? "Multiple Days"
    : "Full Day";  
  const rawDays = await countWorkingDays(intent.date, intent.end_date);
  const daysCount =
  intent.duration === "morning" || intent.duration === "afternoon"
    ? 0.5
    : rawDays;  
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

  // Send preview card and capture the activity ID for reliable edit targeting
  const sentActivity = await ctx.send(buildPreviewCard({
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

  // FIX: store previewCardActivityId alongside pending data.
  // This is the key that makes edit card replacement reliable — even when
  // Teams misroutes the Edit click through the message handler where
  // activity.replyToId may not point to the correct card.
  await savePendingRequest({
    userId:               ctx.userId,
    userName:             ctx.userName,
    intent:               intent.intent,
    date:                 intent.date,
    end_date:             intent.end_date,
    duration:             intent.duration,
    days_count:           daysCount,
    reason:               intent.reason,
    balanceResult,
    history:              [],
    previewCardActivityId: sentActivity?.id ?? null,  // ← stored here
  });
}

// ── Submit Request ────────────────────────────────────────────────────────

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

  // FIX: guard against stale pending from a previous session being submitted.
  // If the pending data has no intent or date, it is corrupt — clear and abort.
  if (!pending.intent || !pending.date) {
    await clearPendingRequest(userId);
    await send("Your request data appears incomplete. Please start a new request.");
    return;
  }

  const employee = await findEmployee(userName);
  if (!employee) { await send(`Employee not found. Please ask HR to add you.`); await clearPendingRequest(userId); return; }

  const { intent, date, end_date, duration, days_count, reason, balanceResult } = pending;
  const displayDate     = formatDisplayDate(date);
  const displayEndDate  = end_date ? formatDisplayDate(end_date) : null;
  const durationLabel   = duration === "half_day" ? "Half Day" : duration === "multi_day" ? "Multiple Days" : "Full Day";
  const isTeamLead      = employee.role === "teamlead";
  const approverName    = isTeamLead ? employee.manager          : employee.teamlead;
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
        const emp = allEmps.find((e: any) => e.name.toLowerCase() === r.employee.toLowerCase());
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