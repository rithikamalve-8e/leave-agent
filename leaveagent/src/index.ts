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
  });

  const cmd = userMessage.toLowerCase();

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

  await send("Processing your request...");

  const intent = await parseLeaveIntent(userMessage);
  console.log(`[LeaveAgent] Intent:`, JSON.stringify(intent));

  if (intent.needs_clarification || intent.intent === "UNKNOWN") {
    await send(
      intent.clarification_question ??
      "Could you rephrase? Try: 'WFH tomorrow', 'Sick today', or 'Leave on Friday'."
    );
    return;
  }

  if (isDuplicateRequest(userName, intent.date)) {
    await send(
      `You already have a request for ${formatDisplayDate(intent.date)}. Type 'my requests' to view it.`
    );
    return;
  }

  const employee = findEmployee(userName);
  if (!employee) {
    await send(
      `I couldn't find ${userName} in the employee directory. Please ask HR to add you.`
    );
    return;
  }

  const displayDate   = formatDisplayDate(intent.date);
  const durationLabel = intent.duration === "half_day" ? "Half Day" : "Full Day";

  addLeaveRequest({
    employee:     userName,
    email:        employee.email,
    type:         intent.intent,
    date:         intent.date,
    end_date:     intent.end_date ?? "",
    duration:     intent.duration,
    status:       "Pending",
    requested_at: new Date().toISOString(),
  });

  await send(buildConfirmationCard(userName, intent.intent, displayDate, durationLabel));

  // Build approval card payload once — used in both inline and proactive paths
  const approvalCardPayload = buildApprovalCardContent({
    employeeName:  userName,
    employeeEmail: employee.email,
    requestType:   intent.intent,
    date:          intent.date,
    displayDate,
    duration:      durationLabel,
  });

  const managerRef        = getConversationRef(employee.manager_teams_id ?? "");
  const managerIsSameUser = employee.manager_teams_id === userId; // devtools: both are "devtools"
  const managerHasRef     = !!(managerRef?.conversationId);

  if (managerIsSameUser || !managerHasRef) {
    // DEVTOOLS / DEMO: show approval card inline so Approve/Reject buttons are clickable
    await send("---- Manager Approval Request ----");
    await send({
      type: "message",
      attachments: [{
        contentType: "application/vnd.microsoft.card.adaptive",
        content: approvalCardPayload,
      }],
    } as any);
  } else {
    // PRODUCTION: send approval card proactively to manager's own conversation
    try {
      await api.conversations.activities(managerRef!.conversationId).create({
        type: "message",
        attachments: [{
          contentType: "application/vnd.microsoft.card.adaptive",
          content: approvalCardPayload,
        }],
      } as any);
      console.log(`[LeaveAgent] Approval card sent to manager ${employee.manager}`);
    } catch (err) {
      // Fallback: show inline if proactive send fails
      console.warn(`[LeaveAgent] Proactive failed, showing inline:`, err);
      await send("---- Manager Approval Request ----");
      await send({
        type: "message",
        attachments: [{
          contentType: "application/vnd.microsoft.card.adaptive",
          content: approvalCardPayload,
        }],
      } as any);
    }
  }
}); // <-- closes app.on("message")

// ─────────────────────────────────────────────
// Adaptive Card Actions — Manager clicks Approve / Reject
// ─────────────────────────────────────────────

app.on("card.action", async ({ activity, send, api }) => {
  const data        = (activity.value as any)?.action?.data;
  const managerName = activity.from.name ?? "Manager";

  console.log(`[LeaveAgent] card.action from ${managerName}:`, data);

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

  if (!updated) {
    return {
      statusCode: 200,
      type: "application/vnd.microsoft.card.adaptive",
      value: buildAlreadyProcessedCardContent(),
    };
  }

  const employee = findEmployee(employeeName);
  const cardData  = { employeeName, requestType, date, displayDate };

  // Notify employee — inline in devtools, proactive DM in production
  const employeeRef        = getConversationRef(employee?.teams_id ?? "");
  const employeeIsSameConv = !employeeRef?.conversationId ||
    employeeRef.conversationId === activity.conversation.id;

  if (employeeIsSameConv) {
    // DEVTOOLS: show employee status card inline
    await send(`---- Employee Notification: Request ${status} ----`);
    await send({
      type: "message",
      attachments: [{
        contentType: "application/vnd.microsoft.card.adaptive",
        content: buildStatusCardContent(requestType, displayDate, status, managerName),
      }],
    } as any);
  } else {
    // PRODUCTION: DM employee proactively
    try {
      await api.conversations.activities(employeeRef!.conversationId).create({
        type: "message",
        attachments: [{
          contentType: "application/vnd.microsoft.card.adaptive",
          content: buildStatusCardContent(requestType, displayDate, status, managerName),
        }],
      } as any);
      console.log(`[LeaveAgent] DMed employee ${employeeName}`);
    } catch (err) {
      console.warn(`[LeaveAgent] Could not DM employee:`, err);
    }
  }

  if (status === "Approved" && employee) {
    await send(buildAnnouncementCard({
      employee:     employeeName,
      email:        employee.email,
      type:         requestType,
      date,
      duration:     "full_day",
      status:       "Approved",
      approved_by:  managerName,
      requested_at: new Date().toISOString(),
    }));
  }

  return {
    statusCode: 200,
    type: "application/vnd.microsoft.card.adaptive",
    value: status === "Approved"
      ? buildApprovedCardContent(cardData, managerName)
      : buildRejectedCardContent(cardData, managerName),
  };
}); // <-- closes app.on("card.action")

// ─────────────────────────────────────────────
// Bot installed — welcome message
// ─────────────────────────────────────────────

app.on("install.add", async ({ send }) => {
  await send(
    "Hi! I'm LeaveAgent, your AI-powered leave assistant.\n\n" +
    "Try: 'WFH tomorrow', 'Sick today', 'Leave on Friday'\n\n" +
    "Type 'help' for all commands."
  );
});

// ─────────────────────────────────────────────
// Start
// ─────────────────────────────────────────────

(async () => {
  await app.start(+(process.env.PORT ?? 3978));
  console.log(`\nLeaveAgent running on port ${process.env.PORT ?? 3978}`);
  console.log(`Endpoint: http://localhost:${process.env.PORT ?? 3978}/api/messages\n`);
})();