/**
 * notificationService.ts
 *
 * Central service for all outbound bot notifications:
 * - Proactive DMs to employees, approvers, HR
 * - Group channel announcements
 * - OAuth token fetching for scheduler-triggered messages
 *
 * Two modes of sending:
 * 1. api-mode  — inside a message/card.action handler where `api` is available
 * 2. token-mode — inside schedulers where there is no active request context,
 *                 so we fetch a fresh OAuth token and call Bot Framework REST directly
 */

import {
  buildStatusCardContent,
  buildApprovalCardContent,
  buildAnnouncementCard,
  buildHRAlertCard,
  buildApproverReminderCard,
  buildHRTakeoverCard,
  buildBalanceAdjustedCard,
  buildDeletedNotificationCard,
  buildHolidayAnnouncementCard,
  formatDisplayDate,
  getTypeEmoji,
  getTypeLabel,
  LeaveRecord,
} from "./cards";

import { getConversationRef, getAllConversationRefs } from "./postgresManager";

// ── Types ──────────────────────────────────────────────────────────────────

export interface NotificationContext {
  api:   any;    // Teams SDK api object from handler context
  botId: string; // activity.recipient?.id
}

// ── OAuth Token (for scheduler use) ───────────────────────────────────────

export async function getBotToken(): Promise<string> {
  const clientId     = process.env.MICROSOFT_APP_ID      ?? process.env.BOT_ID       ?? "";
  const clientSecret = process.env.MICROSOFT_APP_PASSWORD ?? process.env.BOT_PASSWORD ?? "";
  const tenantId     = process.env.MICROSOFT_APP_TENANT_ID ?? "botframework.com";

  const body = new URLSearchParams({
    grant_type:    "client_credentials",
    client_id:     clientId,
    client_secret: clientSecret,
    scope:         "https://api.botframework.com/.default",
  });

  const res  = await fetch(
    `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`,
    { method: "POST", body }
  );
  const json = await res.json() as any;
  return json.access_token ?? "";
}

// ── Core: send a card to a specific conversation via api ───────────────────

async function sendCardToConversation(
  api:            any,
  conversationId: string,
  botId:          string,
  recipientId:    string,
  cardContent:    any
): Promise<void> {
  await api.conversations.activities(conversationId).create({
    type:         "message",
    from:         { id: botId },
    conversation: { id: conversationId },
    recipient:    { id: recipientId },
    attachments:  [{
      contentType: "application/vnd.microsoft.card.adaptive",
      content:     cardContent,
    }],
  } as any);
}

// ── Core: send plain text to a conversation via api ────────────────────────

async function sendTextToConversation(
  api:            any,
  conversationId: string,
  botId:          string,
  text:           string
): Promise<void> {
  await api.conversations.activities(conversationId).create({
    type:         "message",
    text,
    from:         { id: botId },
    conversation: { id: conversationId },
  } as any);
}

// ── Core: send text via Bot Framework REST (for scheduler use) ─────────────

async function sendTextViaRest(
  conversationId: string,
  text:           string,
  serviceUrl?:    string
): Promise<void> {
  const url      = `${serviceUrl ?? process.env.BOT_SERVICE_URL ?? "https://smba.trafficmanager.net/in/"}v3/conversations/${conversationId}/activities`;
  const token    = await getBotToken();
  const res      = await fetch(url, {
    method:  "POST",
    headers: { "Content-Type": "application/json", "Authorization": `Bearer ${token}` },
    body:    JSON.stringify({ type: "message", text }),
  });

  if (!res.ok) {
    throw new Error(`REST send failed: ${res.status} ${await res.text()}`);
  }
}

// ── Core: send card via Bot Framework REST (for scheduler use) ─────────────

async function sendCardViaRest(
  conversationId: string,
  cardContent:    any,
  serviceUrl?:    string
): Promise<void> {
  const url   = `${serviceUrl ?? process.env.BOT_SERVICE_URL ?? "https://smba.trafficmanager.net/in/"}v3/conversations/${conversationId}/activities`;
  const token = await getBotToken();
  const res   = await fetch(url, {
    method:  "POST",
    headers: { "Content-Type": "application/json", "Authorization": `Bearer ${token}` },
    body:    JSON.stringify({
      type:        "message",
      attachments: [{ contentType: "application/vnd.microsoft.card.adaptive", content: cardContent }],
    }),
  });

  if (!res.ok) {
    throw new Error(`REST card send failed: ${res.status} ${await res.text()}`);
  }
}

// ── Announcement Channel ───────────────────────────────────────────────────

/**
 * Sends a plain text message to the configured announcement group channel.
 * Uses api-mode (inside request handler).
 */
export async function sendAnnouncementText(
  ctx:     NotificationContext,
  message: string
): Promise<void> {
  const convId = process.env.ANNOUNCEMENT_CHANNEL_ID;
  if (!convId) {
    console.warn("[Notify] ANNOUNCEMENT_CHANNEL_ID not set — skipping announcement");
    return;
  }
  try {
    await sendTextToConversation(ctx.api, convId, ctx.botId, message);
    console.log(`[Notify] Announcement sent: ${message}`);
  } catch (err) {
    console.warn("[Notify] Could not send announcement:", err);
  }
}

/**
 * Sends a plain text message to the announcement channel via REST (scheduler use).
 */
export async function sendAnnouncementTextRest(message: string): Promise<void> {
  const convId = process.env.ANNOUNCEMENT_CHANNEL_ID;
  if (!convId) return;
  try {
    await sendTextViaRest(convId, message);
    console.log(`[Notify] Announcement (REST) sent: ${message}`);
  } catch (err) {
    console.warn("[Notify] Could not send announcement via REST:", err);
  }
}

// ── Approval Card → Approver ───────────────────────────────────────────────

/**
 * Sends the approval card to the approver's personal DM.
 * Falls back to inline send() if no conversation ref found.
 */
export async function sendApprovalCardToApprover(
  ctx:              NotificationContext,
  approverTeamsId:  string,
  approverName:     string,
  cardPayload:      any,
  fallbackSend:     Function
): Promise<void> {
  const ref = await getConversationRef(approverTeamsId);

  if (!ref?.conversationId) {
    console.warn(`[Notify] No ref for approver ${approverName} — sending inline`);
    await fallbackSend(`---- Approval Request for ${approverName} ----`);
    await fallbackSend({
      type: "message",
      attachments: [{ contentType: "application/vnd.microsoft.card.adaptive", content: cardPayload }],
    } as any);
    return;
  }

  try {
    await sendCardToConversation(ctx.api, ref.conversationId, ref.botId, approverTeamsId, cardPayload);
    console.log(`[Notify] Approval card sent to ${approverName}`);
  } catch (err) {
    console.warn(`[Notify] Proactive to approver failed, sending inline:`, err);
    await fallbackSend(`---- Approval Request for ${approverName} ----`);
    await fallbackSend({
      type: "message",
      attachments: [{ contentType: "application/vnd.microsoft.card.adaptive", content: cardPayload }],
    } as any);
  }
}

// ── Status Card → Employee ─────────────────────────────────────────────────

/**
 * Sends the approved/rejected status card to the employee's personal DM.
 * Falls back to inline if same conversation or no ref.
 */
export async function sendStatusCardToEmployee(
  ctx:              NotificationContext,
  employeeTeamsId:  string,
  approverTeamsId:  string,
  currentConvId:    string,
  requestType:      string,
  displayDate:      string,
  status:           "Approved" | "Rejected",
  approverName:     string,
  rejectionReason?: string,
  fallbackSend?:    Function
): Promise<void> {
  const ref          = await getConversationRef(employeeTeamsId);
  const cardContent  = buildStatusCardContent(requestType, displayDate, status, approverName, rejectionReason);
  const isSameUser   = employeeTeamsId === approverTeamsId;
  const hasDifferentConv = ref?.conversationId && ref.conversationId !== currentConvId;

  if (!isSameUser && hasDifferentConv) {
    try {
      await sendCardToConversation(ctx.api, ref!.conversationId, ref!.botId, employeeTeamsId, cardContent);
      console.log(`[Notify] Status card DMed to employee`);
      return;
    } catch (err) {
      console.warn("[Notify] Could not DM employee:", err);
    }
  }

  // Fallback — send in current conversation
  if (fallbackSend) {
    await fallbackSend({
      type: "message",
      attachments: [{ contentType: "application/vnd.microsoft.card.adaptive", content: cardContent }],
    } as any);
  }
}

// ── Workforce Availability Card → Manager ─────────────────────────────────

/**
 * Sends the Workforce Availability card to the manager's DM after approval.
 * Skips if manager is the same as the approver (no duplicate).
 */
export async function sendWorkforceCardToManager(
  ctx:             NotificationContext,
  employee:        any,
  employeeName:    string,
  requestType:     string,
  date:            string,
  leaveRecord:     any,
  actualDuration:  string,
  actualDaysCount: number,
  approverName:    string,
  approverTeamsId: string
): Promise<void> {
  const managerTeamsId = employee.role === "teamlead"
    ? employee.manager_teams_id
    : employee.teamlead_teams_id;

  if (!managerTeamsId || managerTeamsId === approverTeamsId) return;

  const ref = await getConversationRef(managerTeamsId);
  if (!ref?.conversationId) {
    console.warn(`[Notify] No ref for manager — they need to DM the bot first`);
    return;
  }

  const cardContent = buildAnnouncementCard({
    employee:     employeeName,
    email:        employee.email,
    type:         requestType,
    date,
    end_date:     leaveRecord?.end_date,
    duration:     actualDuration,
    days_count:   actualDaysCount,
    lop_days:     leaveRecord?.lop_days ?? 0,
    status:       "Approved",
    approved_by:  approverName,
    requested_at: new Date().toISOString(),
  }).attachments[0].content;

  try {
    await sendCardToConversation(ctx.api, ref.conversationId, ref.botId, managerTeamsId, cardContent);
    console.log(`[Notify] Workforce card sent to manager`);
  } catch (err) {
    console.warn("[Notify] Could not send workforce card to manager:", err);
  }
}

// ── HR Alert ───────────────────────────────────────────────────────────────

/**
 * Sends an alert card to HR on leave request submission or approval.
 * HR_TEAMS_ID must be set in env.
 */
export async function sendHRAlert(
  ctx:          NotificationContext,
  event:        "submitted" | "approved" | "rejected",
  employeeName: string,
  requestType:  string,
  displayDate:  string,
  actionBy?:    string,
  reason?:      string
): Promise<void> {
  const hrTeamsId = process.env.HR_TEAMS_ID;
  if (!hrTeamsId) {
    console.warn("[Notify] HR_TEAMS_ID not set — skipping HR alert");
    return;
  }

  const ref = await getConversationRef(hrTeamsId);
  if (!ref?.conversationId) {
    console.warn("[Notify] No conversation ref for HR — they need to DM the bot first");
    return;
  }

  const cardContent = buildHRAlertCard(event, employeeName, requestType, displayDate, actionBy, reason)
    .attachments[0].content;

  try {
    await sendCardToConversation(ctx.api, ref.conversationId, ref.botId, hrTeamsId, cardContent);
    console.log(`[Notify] HR alert sent: ${event} — ${employeeName}`);
  } catch (err) {
    console.warn("[Notify] Could not send HR alert:", err);
  }
}

// ── Announcement after approval ────────────────────────────────────────────

/**
 * Builds and sends the plain-text announcement to the group channel after approval.
 * e.g. "📅 Rithika MR will be working from home today."
 */
export async function sendApprovalAnnouncement(
  ctx:          NotificationContext,
  employeeName: string,
  requestType:  string,
  date:         string,
  displayDate:  string,
  endDate?:     string | null
): Promise<void> {
  const typePhrase =
    requestType === "WFH"       ? "working from home"   :
    requestType === "SICK"      ? "on sick leave"        :
    requestType === "LEAVE"     ? "on planned leave"     :
    requestType === "MATERNITY" ? "on maternity leave"   :
    requestType === "PATERNITY" ? "on paternity leave"   :
    requestType === "MARRIAGE"  ? "on marriage leave"    :
    requestType === "ADOPTION"  ? "on adoption leave"    : "unavailable";

  const todayStr   = new Date().toISOString().split("T")[0];
  const datePhrase = date === todayStr ? "today" : `on ${displayDate}`;

  const message = endDate && endDate !== date
    ? `📅 ${employeeName} will be ${typePhrase} from ${displayDate} to ${formatDisplayDate(endDate)}.`
    : `📅 ${employeeName} will be ${typePhrase} ${datePhrase}.`;

  await sendAnnouncementText(ctx, message);
}

// ── Balance Adjustment Notification → Employee ─────────────────────────────

/**
 * Notifies the employee when HR adjusts their leave balance.
 */
export async function sendBalanceAdjustedNotification(
  ctx:          NotificationContext,
  employeeTeamsId: string,
  employeeName: string,
  adjustment:   number,
  newBalance:   number,
  reason:       string,
  adjustedBy:   string
): Promise<void> {
  const ref = await getConversationRef(employeeTeamsId);
  if (!ref?.conversationId) {
    console.warn(`[Notify] No ref for ${employeeName} — cannot notify balance adjustment`);
    return;
  }

  const cardContent = buildBalanceAdjustedCard(employeeName, adjustment, newBalance, reason, adjustedBy)
    .attachments[0].content;

  try {
    await sendCardToConversation(ctx.api, ref.conversationId, ref.botId, employeeTeamsId, cardContent);
    console.log(`[Notify] Balance adjustment notification sent to ${employeeName}`);
  } catch (err) {
    console.warn("[Notify] Could not notify employee of balance adjustment:", err);
  }
}

// ── Delete Notification → Employee + Approver ──────────────────────────────

/**
 * Notifies both the employee and their approver when HR deletes a request.
 */
export async function sendDeleteNotifications(
  ctx:             NotificationContext,
  employeeTeamsId: string,
  approverTeamsId: string,
  employeeName:    string,
  requestType:     string,
  displayDate:     string,
  deletedBy:       string,
  reason:          string
): Promise<void> {
  const cardContent = buildDeletedNotificationCard(employeeName, requestType, displayDate, deletedBy, reason)
    .attachments[0].content;

  // Notify employee
  const empRef = await getConversationRef(employeeTeamsId);
  if (empRef?.conversationId) {
    try {
      await sendCardToConversation(ctx.api, empRef.conversationId, empRef.botId, employeeTeamsId, cardContent);
      console.log(`[Notify] Delete notification sent to ${employeeName}`);
    } catch (err) {
      console.warn("[Notify] Could not notify employee of deletion:", err);
    }
  }

  // Notify approver (only if different from employee)
  if (approverTeamsId && approverTeamsId !== employeeTeamsId) {
    const aprRef = await getConversationRef(approverTeamsId);
    if (aprRef?.conversationId) {
      try {
        await sendCardToConversation(ctx.api, aprRef.conversationId, aprRef.botId, approverTeamsId, cardContent);
        console.log(`[Notify] Delete notification sent to approver`);
      } catch (err) {
        console.warn("[Notify] Could not notify approver of deletion:", err);
      }
    }
  }
}

// ── Holiday Notification → All Employees ──────────────────────────────────

/**
 * Sends a holiday announcement card to all registered employees via DM,
 * and posts a text message to the announcement channel.
 */
export async function sendHolidayNotificationToAll(
  ctx:    NotificationContext,
  date:   string,
  name:   string,
  addedBy: string,
  action: "added" | "edited" | "rescheduled" | "deleted" = "added"
): Promise<void> {
  const announcementText =
    action === "deleted"
      ? `🗑️ Holiday removed: ${name} (${formatDisplayDate(date)}) by ${addedBy}.`
      : action === "rescheduled"
      ? `📅 Holiday rescheduled: ${name} moved to ${formatDisplayDate(date)} by ${addedBy}.`
      : action === "edited"
      ? `✏️ Holiday updated: ${name} on ${formatDisplayDate(date)} by ${addedBy}.`
      : `🎉 New holiday added: ${name} on ${formatDisplayDate(date)} by ${addedBy}.`;

  // Post to group channel
  await sendAnnouncementText(ctx, announcementText);

  // DM all registered employees
  const cardContent = buildHolidayAnnouncementCard(date, name, addedBy, action)
    .attachments[0].content;

  const allRefs = await getAllConversationRefs();
  let   sent    = 0;

  for (const ref of allRefs) {
    if (!ref.conversationId || !ref.conversationId.startsWith("a:")) continue; // personal DMs only
    try {
      await sendCardToConversation(ctx.api, ref.conversationId, ref.botId, ref.userId, cardContent);
      sent++;
    } catch (err) {
      console.warn(`[Notify] Could not DM ${ref.userName} for holiday:`, err);
    }
  }

  console.log(`[Notify] Holiday notification sent to ${sent} employees`);
}

// ── Month-End: Approver Reminder ───────────────────────────────────────────

/**
 * Sends month-end reminder cards to each approver with their pending requests.
 * Uses REST (called from scheduler, no api context).
 */
export async function sendApproverReminders(
  approverGroups: { approverName: string; approverTeamsId: string; records: LeaveRecord[] }[],
  month:          string
): Promise<void> {
  for (const group of approverGroups) {
    if (group.records.length === 0) continue;

    const ref = await getConversationRef(group.approverTeamsId);
    if (!ref?.conversationId) {
      console.warn(`[Notify] No ref for approver ${group.approverName} — skipping reminder`);
      continue;
    }

    const cardContent = buildApproverReminderCard(group.approverName, group.records, month)
      .attachments[0].content;

    try {
      await sendCardViaRest(ref.conversationId, cardContent, ref.serviceUrl);
      console.log(`[Notify] Month-end reminder sent to ${group.approverName}`);
    } catch (err) {
      console.warn(`[Notify] Could not send reminder to ${group.approverName}:`, err);
    }
  }
}

// ── Month-End: HR Takeover ─────────────────────────────────────────────────

/**
 * Sends unactioned request cards to HR at end of month.
 * Uses REST (called from scheduler, no api context).
 */
export async function sendHRTakeover(
  records: LeaveRecord[],
  month:   string
): Promise<void> {
  if (records.length === 0) return;

  const hrTeamsId = process.env.HR_TEAMS_ID;
  if (!hrTeamsId) {
    console.warn("[Notify] HR_TEAMS_ID not set — cannot send HR takeover");
    return;
  }

  const ref = await getConversationRef(hrTeamsId);
  if (!ref?.conversationId) {
    console.warn("[Notify] No conversation ref for HR — cannot send takeover");
    return;
  }

  const cardContent = buildHRTakeoverCard(records, month).attachments[0].content;

  try {
    await sendCardViaRest(ref.conversationId, cardContent, ref.serviceUrl);
    console.log(`[Notify] HR takeover card sent for ${records.length} unactioned requests`);
  } catch (err) {
    console.warn("[Notify] Could not send HR takeover:", err);
  }
}

// ── Daily Summary (Scheduler) ──────────────────────────────────────────────

/**
 * Sends the daily workforce summary text to the announcement channel.
 * Called from scheduler — uses REST mode.
 */
export async function sendDailySummaryRest(records: LeaveRecord[]): Promise<void> {
  if (records.length === 0) return;

  const wfh   = records.filter((r) => r.type === "WFH").map((r) => r.employee);
  const leave = records.filter((r) => r.type !== "WFH").map((r) => r.employee);

  const dateLabel = new Date().toLocaleDateString("en-GB", { day: "numeric", month: "long" });
  let   msg       = `📋 Workforce Availability – ${dateLabel}:`;
  if (wfh.length)                  msg += ` WFH: ${wfh.join(", ")}`;
  if (wfh.length && leave.length)  msg += " |";
  if (leave.length)                msg += ` Leave: ${leave.join(", ")}`;

  await sendAnnouncementTextRest(msg);
}
