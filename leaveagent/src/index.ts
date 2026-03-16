import dotenv from "dotenv";
dotenv.config({ path: "env/.env.local" });
 
import { App } from "@microsoft/teams.apps";
import { DevtoolsPlugin } from "@microsoft/teams.dev";
import { parseLeaveIntent } from "./app/groqParser";
import {
  findEmployee,
  addLeaveRequest,
  updateLeaveStatus,
  isDuplicateRequest,
  getTodaysAbsences,
  getAllLeaveRequests,
  LeaveRecord,
  saveConversationRef,
  getConversationRef,
  countWorkingDays,
  checkLeaveBalance,
} from "./app/excelManager";
import {
  buildApprovalCardContent,
  buildApprovedCardContent,
  buildRejectedCardContent,
  buildConfirmationCard,
  buildStatusCardContent,
  buildDailySummaryCard,
  buildAnnouncementCard,
  buildHelpCard,
  buildMyRequestsCard,
  buildAlreadyProcessedCardContent,
  formatDisplayDate,
} from "./app/cards";
 
const app = new App({
  clientId:     process.env.MICROSOFT_APP_ID       ?? process.env.BOT_ID ?? "",
  clientSecret: process.env.MICROSOFT_APP_PASSWORD ?? process.env.BOT_PASSWORD ?? "",
  tenantId:     process.env.MICROSOFT_APP_TENANT_ID ?? "",
  plugins: [new DevtoolsPlugin()],
});

// ── Helper: send text to announcement group channel ────────────────────────

async function sendAnnouncement(api: any, botId: string, message: string) {
  const announcementConvId = process.env.ANNOUNCEMENT_CHANNEL_ID;
  if (!announcementConvId) return;
  try {
    await api.conversations.activities(announcementConvId).create({
      type:         "message",
      text:         message,
      from:         { id: botId },
      conversation: { id: announcementConvId },
    } as any);
    console.log(`[LeaveAgent] Announcement sent: ${message}`);
  } catch (err) {
    console.warn(`[LeaveAgent] Could not send announcement:`, err);
  }
}

// ── Helper: send Workforce Availability card to manager's DM ──────────────

async function sendManagerCard(
  api: any,
  employee: any,
  employeeName: string,
  requestType: string,
  date: string,
  leaveRecord: any,
  actualDuration: string,
  actualDaysCount: number,
  approverName: string,
  approverTeamsId: string
) {
  const managerTeamsId = employee.role === "teamlead"
    ? employee.manager_teams_id
    : employee.teamlead_teams_id;

  // Don't send to the same person who just approved
  if (!managerTeamsId || managerTeamsId === approverTeamsId) return;

  const managerRef = getConversationRef(managerTeamsId);
  if (!managerRef?.conversationId) {
    console.warn(`[LeaveAgent] No conversation ref for manager — they need to DM the bot first.`);
    return;
  }

  try {
    await api.conversations.activities(managerRef.conversationId).create({
      type:         "message",
      from:         { id: managerRef.botId },
      conversation: { id: managerRef.conversationId },
      recipient:    { id: managerTeamsId },
      attachments:  [{
        contentType: "application/vnd.microsoft.card.adaptive",
        content:     buildAnnouncementCard({
          employee:     employeeName,
          email:        employee.email,
          type:         requestType,
          date,
          end_date:     leaveRecord?.end_date,
          duration:     actualDuration,
          days_count:   actualDaysCount,
          status:       "Approved",
          approved_by:  approverName,
          requested_at: new Date().toISOString(),
        }).attachments[0].content,
      }],
    } as any);
    console.log(`[LeaveAgent] Workforce availability card sent to manager`);
  } catch (err) {
    console.warn(`[LeaveAgent] Could not send card to manager:`, err);
  }
}

// ── Helper: build type label ───────────────────────────────────────────────

function getTypeLabel(requestType: string): string {
  return requestType === "WFH"       ? "working from home" :
         requestType === "SICK"      ? "on sick leave" :
         requestType === "LEAVE"     ? "on planned leave" :
         requestType === "MATERNITY" ? "on maternity leave" :
         requestType === "PATERNITY" ? "on paternity leave" :
         requestType === "MARRIAGE"  ? "on marriage leave" :
         requestType === "ADOPTION"  ? "on adoption leave" : "unavailable";
}

// ── Incoming Messages ──────────────────────────────────────────────────────
 
app.on("message", async ({ activity, send, api }) => {
  const userMessage = (activity.text ?? "").replace(/<[^>]+>/g, "").trim();
  const userId      = activity.from.id;
  const userName    = activity.from.name ?? "Employee";
 
  console.log(`[LeaveAgent] "${userMessage}" from ${userName} (${userId})`);
 
  saveConversationRef(userId, {
    userId,
    userName,
    conversationId: activity.conversation.id,
    serviceUrl:     activity.serviceUrl,
    tenantId:       activity.conversation.tenantId,
    botId:          activity.recipient?.id ?? "bot",
  }, activity.conversation.id.startsWith("a:")); // true = personal DM

  // ── Intercept card button clicks misrouted as messages ─────────────────
  const activityValue = (activity as any).value;
  if (activityValue?.action === "approve" || activityValue?.action === "reject") {
    const data        = activityValue;
    const approverName = userName;
    const { action, employeeName, date, requestType = "WFH" } = data;
    const status: "Approved" | "Rejected" = action === "approve" ? "Approved" : "Rejected";
    const displayDate = formatDisplayDate(date);

    const updated = updateLeaveStatus(employeeName, date, status, approverName);
    if (!updated) {
      await send("This request has already been processed.");
      return;
    }

    const employee    = findEmployee(employeeName);
    const allRecords  = getAllLeaveRequests();
    const leaveRecord = allRecords.find(
      (r) => r.employee?.toLowerCase() === employeeName.toLowerCase() && r.date === date
    );
    const actualDuration  = leaveRecord?.duration  ?? "full_day";
    const actualDaysCount = leaveRecord?.days_count ?? 1;

    await send(`Request ${status} for ${employeeName} on ${displayDate}.`);

    // Notify employee via DM
    const employeeRef = getConversationRef(employee?.teams_id ?? "");
    console.log(`[DEBUG] employeeRef:`, JSON.stringify(employeeRef));
    console.log(`[DEBUG] activity.conversation.id:`, activity.conversation.id);
    console.log(`[DEBUG] employee?.teams_id:`, employee?.teams_id);

    if (employeeRef?.conversationId && employee?.teams_id !== activity.from.id) {
      try {
        await api.conversations.activities(employeeRef.conversationId).create({
          type:         "message",
          from:         { id: employeeRef.botId },
          conversation: { id: employeeRef.conversationId },
          recipient:    { id: employee?.teams_id },
          attachments:  [{ contentType: "application/vnd.microsoft.card.adaptive", content: buildStatusCardContent(requestType, displayDate, status, approverName) }],
        } as any);
      } catch (err) {
        console.warn(`[LeaveAgent] Could not DM employee:`, err);
      }
    } else {
      await send({
        type: "message",
        attachments: [{ contentType: "application/vnd.microsoft.card.adaptive", content: buildStatusCardContent(requestType, displayDate, status, approverName) }],
      } as any);
    }

    if (status === "Approved" && employee) {
      const todayStr   = new Date().toISOString().split("T")[0];
      const datePhrase = date === todayStr ? "today" : `on ${displayDate}`;
      const message    = leaveRecord?.end_date && leaveRecord.end_date !== date
        ? `📅 ${employeeName} will be ${getTypeLabel(requestType)} from ${displayDate} to ${formatDisplayDate(leaveRecord.end_date)}.`
        : `📅 ${employeeName} will be ${getTypeLabel(requestType)} ${datePhrase}.`;

      // Post plain text to group channel
      await sendAnnouncement(api, activity.recipient?.id ?? "", message);

      // Send Workforce Availability card to manager's DM
      await sendManagerCard(api, employee, employeeName, requestType, date, leaveRecord, actualDuration, actualDaysCount, approverName, userId);
    }
    return;
  }

  const cmd = userMessage.toLowerCase();
 
  // ── Special commands ───────────────────────────────────────────────
  if (cmd === "help")    { await send(buildHelpCard()); return; }
  if (cmd === "summary") { await send(buildDailySummaryCard(getTodaysAbsences())); return; }

  if (cmd === "my requests") {
    const mine = getAllLeaveRequests()
      .filter((r) => r.employee?.toLowerCase() === userName.toLowerCase())
      .slice(-5);
    await send(buildMyRequestsCard(userName, mine));
    return;
  }
 
  if (
    cmd === "my balance"       ||
    cmd === "leave balance"    ||
    cmd === "my leave balance" ||
    cmd === "balance"          ||
    cmd.includes("how many leave")  ||
    cmd.includes("leaves left")     ||
    cmd.includes("leave balance")   ||
    cmd.includes("how many days")   ||
    cmd.includes("days remaining")  ||
    cmd.includes("days left")
  ) {
    const employee = findEmployee(userName);
    if (!employee) {
      await send(`I couldn't find ${userName} in the employee directory. Please ask HR to add you.`);
      return;
    }
    const allRequests = getAllLeaveRequests();
    const pendingDays = allRequests
      .filter((r) => r.employee?.toLowerCase() === userName.toLowerCase() && r.status === "Pending" && r.type === "LEAVE")
      .reduce((sum, r) => sum + (Number(r.days_count) || 1), 0);
    const available = Math.max(0, (employee.leave_balance ?? 0) - pendingDays);
    await send(
      `Leave Balance for ${userName}\n\n` +
      `Annual Balance: ${employee.leave_balance ?? 0} day(s)\n` +
      `Pending Requests: ${pendingDays} day(s) awaiting approval\n` +
      `Available to Book: ${available} day(s)\n\n` +
      `Note: WFH does not consume your leave balance. All other leave types including Sick Leave are deducted from your 22 annual days.`
    );
    return;
  }
 
  // ── Test approve/reject command ──────────────────────────────────
  const approveMatch = userMessage.match(/^(approve|reject) leave (\S+) (\d{4}-\d{2}-\d{2})$/i);
  if (approveMatch) {
    const action    = approveMatch[1].toLowerCase();
    const empName   = approveMatch[2];
    const leaveDate = approveMatch[3];
    const status    = action === "approve" ? "Approved" : "Rejected";
    const updated   = updateLeaveStatus(empName, leaveDate, status, userName);
    await send(updated
      ? `${status} leave for ${empName} on ${leaveDate} by ${userName}.`
      : `No pending leave found for ${empName} on ${leaveDate}.`
    );
    return;
  }
 
  // ── Block requests on behalf of others ──────────────────────────
 const onBehalfMatch = userMessage.match(/(?:leave|wfh|sick).*\bfor\s+(\w+)/i);
  if (onBehalfMatch && onBehalfMatch[1].toLowerCase() !== userName.toLowerCase()) {
    await send(`You can only submit requests for yourself. Please ask ${onBehalfMatch[1]} to submit their own request.`);
    return;
  }

  // ── AI intent parsing ──────────────────────────────────────────────
  await send("Processing your request...");
  const intent = await parseLeaveIntent(userMessage);
  console.log(`[LeaveAgent] Intent:`, JSON.stringify(intent));
 
  if (intent.needs_clarification || intent.intent === "UNKNOWN") {
    await send(intent.clarification_question ?? "Could you rephrase? Try: 'WFH tomorrow', 'Sick today', or 'Leave from 20th to 25th'.");
    return;
  }
 
  if (isDuplicateRequest(userName, intent.date)) {
    await send(`You already have a request for ${formatDisplayDate(intent.date)}. Type 'my requests' to view it.`);
    return;
  }
 
  const employee = findEmployee(userName);
  if (!employee) {
    await send(`I couldn't find ${userName} in the employee directory. Please ask HR to add you.`);
    return;
  }
 
  const displayDate    = formatDisplayDate(intent.date);
  const displayEndDate = intent.end_date ? formatDisplayDate(intent.end_date) : null;
  const durationLabel  = intent.duration === "half_day" ? "Half Day" : intent.duration === "multi_day" ? "Multiple Days" : "Full Day";
  const isTeamLead     = employee.role === "teamlead";
  const daysCount      = intent.duration === "half_day" ? 0.5 : countWorkingDays(intent.date, intent.end_date);
  const balanceResult  = checkLeaveBalance(employee, daysCount, intent.intent);
 
  if (balanceResult.hasLop) {
    await send(
      `Your current leave balance is ${balanceResult.balance} day(s).\n\n` +
      `You requested ${balanceResult.requested} day(s).\n\n` +
      `${balanceResult.granted} day(s) will be approved from your balance.\n` +
      `${balanceResult.lop} day(s) will be Loss of Pay (LOP) - please contact HR for details.\n\n` +
      `Your request has been submitted and will go to your approver.`
    );
  }
 
  const approverName    = isTeamLead ? employee.manager          : employee.teamlead;
  const approverTeamsId = isTeamLead ? employee.manager_teams_id : employee.teamlead_teams_id;
 
  addLeaveRequest({
    employee:     userName,
    email:        employee.email,
    type:         intent.intent,
    date:         intent.date,
    end_date:     intent.end_date ?? "",
    duration:     intent.duration,
    days_count:   daysCount,
    reason:       intent.reason ?? "",
    status:       "Pending",
    requested_at: new Date().toISOString(),
  });
 
  await send(buildConfirmationCard(userName, intent.intent, displayDate, durationLabel, displayEndDate, daysCount, intent.reason, balanceResult));
 
  const approvalCardPayload = buildApprovalCardContent({
    employeeName:  userName,
    employeeEmail: employee.email,
    requestType:   intent.intent,
    date:          intent.date,
    displayDate,
    duration:      durationLabel,
    endDate:       displayEndDate,
    daysCount,
    reason:        intent.reason,
    balanceResult,
  });
 
  const approverRef        = getConversationRef(approverTeamsId ?? "");
  const approverIsSameUser = approverTeamsId === userId;
  const approverHasRef     = !!(approverRef?.conversationId);
 
  if (approverIsSameUser || !approverHasRef) {
    await send(`---- Approval Request for ${approverName} ----`);
    await send({ type: "message", attachments: [{ contentType: "application/vnd.microsoft.card.adaptive", content: approvalCardPayload }] } as any);
  } else {
    try {
      await api.conversations.activities(approverRef!.conversationId).create({
        type:         "message",
        from:         { id: approverRef!.botId },
        conversation: { id: approverRef!.conversationId },
        recipient:    { id: approverTeamsId },
        attachments:  [{ contentType: "application/vnd.microsoft.card.adaptive", content: approvalCardPayload }],
      } as any);
      console.log(`[LeaveAgent] Approval card sent to ${approverName}`);
    } catch (err) {
      console.warn(`[LeaveAgent] Proactive failed, showing inline:`, err);
      await send(`---- Approval Request for ${approverName} ----`);
      await send({ type: "message", attachments: [{ contentType: "application/vnd.microsoft.card.adaptive", content: approvalCardPayload }] } as any);
    }
  }
});
 
// ── Card Actions ───────────────────────────────────────────────────────────
 
app.on("card.action", async ({ activity, send, api }) => {
  const data        = (activity.value as any)?.action?.data;
  const approverName = activity.from.name ?? "Approver";
  const approverTeamsId = activity.from.id;
 
  console.log(`[LeaveAgent] card.action from ${approverName}:`, data);
 
  if (!data?.action || !data?.employeeName || !data?.date) {
    return { statusCode: 200, type: "application/vnd.microsoft.card.adaptive", value: buildAlreadyProcessedCardContent() as any } as any;
  }
 
  const { action, employeeName, date, requestType = "WFH" } = data;
  const status: "Approved" | "Rejected" = action === "approve" ? "Approved" : "Rejected";
  const displayDate = formatDisplayDate(date);
 
  const updated = updateLeaveStatus(employeeName, date, status, approverName);
  if (!updated) {
    return { statusCode: 200, type: "application/vnd.microsoft.card.adaptive", value: buildAlreadyProcessedCardContent() as any } as any;
  }
 
  const employee  = findEmployee(employeeName);
  const cardData  = { employeeName, requestType, date, displayDate };
  const allRecords  = getAllLeaveRequests();
  const leaveRecord = allRecords.find(
    (r) => r.employee?.toLowerCase() === employeeName.toLowerCase() && r.date === date
  );
  const actualDuration  = leaveRecord?.duration  ?? "full_day";
  const actualDaysCount = leaveRecord?.days_count ?? 1;

  // Notify employee via DM
  const employeeRef        = getConversationRef(employee?.teams_id ?? "");
  const employeeIsSameConv = !employeeRef?.conversationId || employeeRef.conversationId === activity.conversation.id;
 
  if (employeeIsSameConv) {
    await send(`---- Employee Notification: Request ${status} ----`);
    await send({ type: "message", attachments: [{ contentType: "application/vnd.microsoft.card.adaptive", content: buildStatusCardContent(requestType, displayDate, status, approverName) }] } as any);
  } else {
    try {
      await api.conversations.activities(employeeRef!.conversationId).create({
        type:         "message",
        from:         { id: employeeRef!.botId },
        conversation: { id: employeeRef!.conversationId },
        recipient:    { id: employee?.teams_id },
        attachments:  [{ contentType: "application/vnd.microsoft.card.adaptive", content: buildStatusCardContent(requestType, displayDate, status, approverName) }],
      } as any);
      console.log(`[LeaveAgent] DMed employee ${employeeName}`);
    } catch (err) {
      console.warn(`[LeaveAgent] Could not DM employee:`, err);
    }
  }
 
  if (status === "Approved" && employee) {
    const todayStr   = new Date().toISOString().split("T")[0];
    const datePhrase = date === todayStr ? "today" : `on ${displayDate}`;
    const message    = leaveRecord?.end_date && leaveRecord.end_date !== date
      ? ` ${employeeName} will be ${getTypeLabel(requestType)} from ${displayDate} to ${formatDisplayDate(leaveRecord.end_date)}.`
      : ` ${employeeName} will be ${getTypeLabel(requestType)} ${datePhrase}.`;

    // Post plain text to group channel
    await sendAnnouncement(api, activity.recipient?.id ?? "", message);

    // Send Workforce Availability card to manager's DM
    await sendManagerCard(api, employee, employeeName, requestType, date, leaveRecord, actualDuration, actualDaysCount, approverName, approverTeamsId);
  }
 
  return {
    statusCode: 200,
    type: "application/vnd.microsoft.card.adaptive",
    value: (status === "Approved"
      ? buildApprovedCardContent(cardData, approverName)
      : buildRejectedCardContent(cardData, approverName)) as any,
  } as any;
});
 
// ── Welcome ────────────────────────────────────────────────────────────────
 
app.on("install.add", async ({ send }) => {
  await send(
    "Hi! I'm LeaveAgent, your AI-powered leave assistant.\n\n" +
    "Try: 'WFH tomorrow', 'Sick today', 'Leave from 20th to 25th'\n\n" +
    "Type 'help' for all commands."
  );
});

// ── Daily Morning Summary ──────────────────────────────────────────────────

function scheduleDailySummary() {
  const announcementConvId = process.env.ANNOUNCEMENT_CHANNEL_ID;
  if (!announcementConvId) {
    console.log("[LeaveAgent] No ANNOUNCEMENT_CHANNEL_ID — daily summary disabled.");
    return;
  }

  function msUntil9am(): number {
    const now     = new Date();
    const next9am = new Date();
    next9am.setHours(9, 0, 0, 0);
    if (now >= next9am) next9am.setDate(next9am.getDate() + 1);
    return next9am.getTime() - now.getTime();
  }

  async function postDailySummary() {
    const today = new Date();
    if (today.getDay() === 0 || today.getDay() === 6) return;

    const todayStr     = today.toISOString().split("T")[0];
    const allRecords   = getAllLeaveRequests();
    const todayRecords = allRecords.filter((r) => r.date === todayStr && r.status === "Approved");
    if (todayRecords.length === 0) return;

    const wfh   = todayRecords.filter((r) => r.type === "WFH").map((r) => r.employee);
    const leave = todayRecords.filter((r) => r.type !== "WFH").map((r) => r.employee);

    const dateLabel = today.toLocaleDateString("en-GB", { day: "numeric", month: "long" });
    let msg = ` Workforce Availability – ${dateLabel}:`;
    if (wfh.length)                  msg += ` WFH: ${wfh.join(", ")}`;
    if (wfh.length && leave.length)  msg += " |";
    if (leave.length)                msg += ` Leave: ${leave.join(", ")}`;

    const serviceUrl = process.env.BOT_SERVICE_URL ?? "https://smba.trafficmanager.net/in/";
    try {
      const token = await getBotToken();
      const res   = await fetch(`${serviceUrl}v3/conversations/${announcementConvId}/activities`, {
        method:  "POST",
        headers: { "Content-Type": "application/json", "Authorization": `Bearer ${token}` },
        body:    JSON.stringify({ type: "message", text: msg }),
      });
      if (res.ok) console.log(`[LeaveAgent] Daily summary sent: ${msg}`);
      else        console.warn(`[LeaveAgent] Daily summary failed: ${res.status}`);
    } catch (err) {
      console.warn("[LeaveAgent] Daily summary error:", err);
    }
  }

  const ms = msUntil9am();
  console.log(`[LeaveAgent] Daily summary scheduled — first run in ${Math.floor(ms/3600000)}h ${Math.floor((ms%3600000)/60000)}m`);
  setTimeout(() => { postDailySummary(); setInterval(postDailySummary, 24 * 60 * 60 * 1000); }, ms);
}

async function getBotToken(): Promise<string> {
  const clientId     = process.env.MICROSOFT_APP_ID      ?? process.env.BOT_ID       ?? "";
  const clientSecret = process.env.MICROSOFT_APP_PASSWORD ?? process.env.BOT_PASSWORD ?? "";
  const tenantId     = process.env.MICROSOFT_APP_TENANT_ID ?? "botframework.com";
  const url  = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;
  const body = new URLSearchParams({
    grant_type: "client_credentials", client_id: clientId,
    client_secret: clientSecret, scope: "https://api.botframework.com/.default",
  });
  const res  = await fetch(url, { method: "POST", body });
  const json = await res.json() as any;
  return json.access_token ?? "";
}
 
// ── Start ──────────────────────────────────────────────────────────────────
 
(async () => {
  await app.start(+(process.env.PORT ?? 3978));
  console.log(`\nLeaveAgent running on port ${process.env.PORT ?? 3978}\n`);
  scheduleDailySummary();
})();