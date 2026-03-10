import dotenv from "dotenv";
dotenv.config({ path: "env/.env.local" });

import { App } from "@microsoft/teams.apps";
import {
  AdaptiveCardActionErrorResponse,
  AdaptiveCardActionMessageResponse,
} from "@microsoft/teams.api";
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

// ─────────────────────────────────────────────
// Initialize Teams App (SDK v2)
// ─────────────────────────────────────────────

const app = new App({
  plugins: [new DevtoolsPlugin()],
});

// ─────────────────────────────────────────────
// Incoming Messages from Employees
// ─────────────────────────────────────────────

app.on("message", async ({ activity, send, api }) => {
  // Strip HTML mention tags Teams adds (e.g. <at>BotName</at>)
  const userMessage = (activity.text ?? "").replace(/<[^>]+>/g, "").trim();
  const userId      = activity.from.id;
  const userName    = activity.from.name ?? "Employee";

  console.log(`[LeaveAgent] "${userMessage}" from ${userName} (${userId})`);

  // Save conversation ref — needed for proactive DMs later
  saveConversationRef(userId, {
    userId,
    userName,
    conversationId: activity.conversation.id,
    serviceUrl:     activity.serviceUrl,
    tenantId:       activity.conversation.tenantId,
    botId:          activity.recipient.id,
  });

  const cmd = userMessage.toLowerCase();

  // ── Special commands ──────────────────────────────────────────────
  if (cmd === "help") {
    await send(buildHelpCard());
    return;
  }

  if (cmd === "summary") {
    await send(buildDailySummaryCard(getTodaysAbsences()));
    return;
  }

  if (cmd === "my requests") {
    const mine = getAllLeaveRequests()
      .filter((r) => r.employee?.toLowerCase() === userName.toLowerCase())
      .slice(-5);
    await send(buildMyRequestsCard(userName, mine));
    return;
  }

  // ── AI intent parsing ─────────────────────────────────────────────
  // Send typing as a plain string activity the SDK accepts
  await send("⏳ Processing your request...");

  const intent = await parseLeaveIntent(userMessage);
  console.log(`[LeaveAgent] Intent:`, JSON.stringify(intent));

  if (intent.needs_clarification || intent.intent === "UNKNOWN") {
    await send(
      intent.clarification_question ??
        "🤔 Could you rephrase? Try: 'WFH tomorrow', 'Sick today', or 'Leave on Friday'."
    );
    return;
  }

  // ── Duplicate check ───────────────────────────────────────────────
  if (isDuplicateRequest(userName, intent.date)) {
    await send(
      `⚠️ You already have a request for **${formatDisplayDate(intent.date)}**.\nType \`my requests\` to view it.`
    );
    return;
  }

  // ── Employee lookup ───────────────────────────────────────────────
  const employee = findEmployee(userName);
  if (!employee) {
    await send(
      `❌ I couldn't find **${userName}** in the employee directory.\nPlease ask HR to add you to the Employees sheet.`
    );
    return;
  }

  const displayDate  = formatDisplayDate(intent.date);
  const durationLabel = intent.duration === "half_day" ? "Half Day" : "Full Day";

  // ── Write to Excel ────────────────────────────────────────────────
  const record: LeaveRecord = {
    employee:     userName,
    email:        employee.email,
    type:         intent.intent,
    date:         intent.date,
    end_date:     intent.end_date ?? "",
    duration:     intent.duration,
    status:       "Pending",
    requested_at: new Date().toISOString(),
  };
  addLeaveRequest(record);

  // ── Confirm to employee ───────────────────────────────────────────
  await send(buildConfirmationCard(userName, intent.intent, displayDate, durationLabel));

  // ── Send approval card to manager (proactive) ─────────────────────
  const managerRef = getConversationRef(employee.manager_teams_id ?? "");
  if (managerRef?.conversationId) {
    try {
      await api.conversations.activities(managerRef.conversationId).create({
        type: "message",
        attachments: [
          {
            contentType: "application/vnd.microsoft.card.adaptive",
            content: buildApprovalCardContent({
              employeeName:  userName,
              employeeEmail: employee.email,
              requestType:   intent.intent,
              date:          intent.date,
              displayDate,
              duration:      durationLabel,
            }),
          },
        ],
      } as any);
      console.log(`[LeaveAgent] ✅ Approval card sent to manager ${employee.manager}`);
    } catch (err) {
      console.warn(`[LeaveAgent] ⚠️ Proactive card to manager failed:`, err);
      await send(`📨 Your manager **${employee.manager}** has been notified.`);
    }
  } else {
    console.warn(`[LeaveAgent] ⚠️ No conversation ref for manager ${employee.manager} — they must message the bot first.`);
    await send(`📨 Your manager **${employee.manager}** has been notified.`);
  }
});

// ─────────────────────────────────────────────
// Adaptive Card Actions — Manager clicks Approve / Reject
// card.action handler must return AdaptiveCardInvokeResponse
// ─────────────────────────────────────────────

app.on("card.action", async ({ activity, send, api }) => {
  const data        = (activity.value as any)?.action?.data;
  const managerName = activity.from.name ?? "Manager";

  console.log(`[LeaveAgent] card.action from ${managerName}:`, data);

  // Validate incoming data
  if (!data?.action || !data?.employeeName || !data?.date) {
    return {
      statusCode: 200,
      type: "application/vnd.microsoft.card.adaptive",
      value: buildAlreadyProcessedCardContent(),
    };
  }

  const { action, employeeName, date, requestType = "WFH" } = data;
  const status: "Approved" | "Rejected" = action === "approve" ? "Approved" : "Rejected";
  const displayDate = formatDisplayDate(date);

  const updated = updateLeaveStatus(employeeName, date, status, managerName);

  // Already processed — return a replaced card
  if (!updated) {
    return {
      statusCode: 200,
      type: "application/vnd.microsoft.card.adaptive",
      value: buildAlreadyProcessedCardContent(),
    };
  }

  const employee = findEmployee(employeeName);
  const cardData = { employeeName, requestType, date, displayDate };

  // ── DM employee with decision ─────────────────────────────────────
  const employeeRef = getConversationRef(employee?.teams_id ?? "");
  if (employeeRef?.conversationId) {
    try {
      await api.conversations.activities(employeeRef.conversationId).create({
        type: "message",
        attachments: [
          {
            contentType: "application/vnd.microsoft.card.adaptive",
            content: buildStatusCardContent(requestType, displayDate, status, managerName),
          },
        ],
      } as any);
      console.log(`[LeaveAgent] ✅ DMed employee ${employeeName}`);
    } catch (err) {
      console.warn(`[LeaveAgent] ⚠️ Could not DM employee:`, err);
    }
  }

  // ── Post workforce announcement if approved ───────────────────────
  if (status === "Approved" && employee) {
    const leaveRecord: LeaveRecord = {
      employee:     employeeName,
      email:        employee.email,
      type:         requestType,
      date,
      duration:     "full_day",
      status:       "Approved",
      approved_by:  managerName,
      requested_at: new Date().toISOString(),
    };
    await send(buildAnnouncementCard(leaveRecord));
  }

  // ── Return updated card — replaces the approval card in Teams ─────
  return {
    statusCode: 200,
    type: "application/vnd.microsoft.card.adaptive",
    value: status === "Approved"
      ? buildApprovedCardContent(cardData, managerName)
      : buildRejectedCardContent(cardData, managerName),
  };
});

// ─────────────────────────────────────────────
// Bot installed — welcome message
// ─────────────────────────────────────────────

app.on("install.add", async ({ send }) => {
  await send(
    "👋 Hi! I'm **LeaveAgent** — your AI-powered leave assistant.\n\n" +
    "Just tell me:\n• `WFH tomorrow`\n• `Sick today`\n• `Leave on Friday`\n\n" +
    "Type `help` for all commands. ✅"
  );
});

// ─────────────────────────────────────────────
// Start
// ─────────────────────────────────────────────

(async () => {
  await app.start(+(process.env.PORT ?? 3978));
  console.log(`\n🤖 LeaveAgent running on port ${process.env.PORT ?? 3978}`);
  console.log(`📡 Endpoint: http://localhost:${process.env.PORT ?? 3978}/api/messages\n`);
})();