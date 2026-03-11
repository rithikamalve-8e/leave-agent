import { IAdaptiveCard } from "@microsoft/teams.cards";
import { LeaveRecord } from "./excelManager";

export interface ApprovalCardData {
  employeeName: string;
  employeeEmail: string;
  requestType: string;
  date: string;
  displayDate: string;
  duration: string;
}

export interface SimpleCardData {
  employeeName: string;
  requestType: string;
  date: string;
  displayDate: string;
}

export interface CardActivity {
  type: "message";
  attachments: Array<{
    contentType: "application/vnd.microsoft.card.adaptive";
    content: IAdaptiveCard;
  }>;
}

export function formatDisplayDate(isoDate: string): string {
  if (!isoDate) return "Unknown date";
  try {
    const d = new Date(isoDate + "T00:00:00");
    return d.toLocaleDateString("en-IN", {
      weekday: "long",
      year: "numeric",
      month: "long",
      day: "numeric",
    });
  } catch {
    return isoDate;
  }
}

export function getTypeLabel(type: string): string {
  switch (type?.toUpperCase()) {
    case "WFH":   return "Work From Home";
    case "LEAVE": return "Planned Leave";
    case "SICK":  return "Sick Leave";
    default:      return "Leave Request";
  }
}

function wrap(content: IAdaptiveCard): CardActivity {
  return {
    type: "message",
    attachments: [{ contentType: "application/vnd.microsoft.card.adaptive", content }],
  };
}

// ── Confirmation Card ──────────────────────────────────────────────────────

export function buildConfirmationCard(
  employeeName: string,
  requestType: string,
  displayDate: string,
  duration: string
): CardActivity {
  return wrap({
    $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
    type: "AdaptiveCard",
    version: "1.4",
    body: [
      {
        type: "TextBlock",
        text: "Request Submitted",
        weight: "Bolder",
        size: "Large",
        color: "Accent",
      },
      {
        type: "FactSet",
        facts: [
          { title: "Employee", value: employeeName },
          { title: "Type",     value: getTypeLabel(requestType) },
          { title: "Date",     value: displayDate },
          { title: "Duration", value: duration },
          { title: "Status",   value: "Awaiting manager approval" },
        ],
      },
      {
        type: "TextBlock",
        text: "Your manager has been notified and will review shortly.",
        wrap: true,
        color: "Good",
        size: "Small",
      },
    ],
  });
}

// ── Approval Card ──────────────────────────────────────────────────────────

export function buildApprovalCardContent(data: ApprovalCardData): IAdaptiveCard {
  return {
    $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
    type: "AdaptiveCard",
    version: "1.4",
    body: [
      {
        type: "TextBlock",
        text: "Leave Approval Required",
        weight: "Bolder",
        size: "Large",
        color: "Warning",
      },
      {
        type: "FactSet",
        facts: [
          { title: "Employee", value: data.employeeName },
          { title: "Email",    value: data.employeeEmail },
          { title: "Type",     value: getTypeLabel(data.requestType) },
          { title: "Date",     value: data.displayDate },
          { title: "Duration", value: data.duration },
        ],
      },
      {
        type: "TextBlock",
        text: "Please approve or reject this request.",
        wrap: true,
        size: "Small",
      },
    ],
    actions: [
      {
        type: "Action.Submit",
        title: "Approve",
        style: "positive",
        data: {
          action: "approve",
          employeeName: data.employeeName,
          date: data.date,
          requestType: data.requestType,
        },
      },
      {
        type: "Action.Submit",
        title: "Reject",
        style: "destructive",
        data: {
          action: "reject",
          employeeName: data.employeeName,
          date: data.date,
          requestType: data.requestType,
        },
      },
    ],
  };
}

// ── Approved Card ──────────────────────────────────────────────────────────

export function buildApprovedCardContent(data: SimpleCardData, approvedBy: string): IAdaptiveCard {
  return {
    $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
    type: "AdaptiveCard",
    version: "1.4",
    body: [
      {
        type: "TextBlock",
        text: "Request Approved",
        weight: "Bolder",
        size: "Large",
        color: "Good",
      },
      {
        type: "FactSet",
        facts: [
          { title: "Employee",    value: data.employeeName },
          { title: "Type",        value: getTypeLabel(data.requestType) },
          { title: "Date",        value: data.displayDate },
          { title: "Approved By", value: approvedBy },
          { title: "Time",        value: new Date().toLocaleTimeString("en-IN") },
        ],
      },
    ],
  };
}

// ── Rejected Card ──────────────────────────────────────────────────────────

export function buildRejectedCardContent(data: SimpleCardData, rejectedBy: string): IAdaptiveCard {
  return {
    $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
    type: "AdaptiveCard",
    version: "1.4",
    body: [
      {
        type: "TextBlock",
        text: "Request Rejected",
        weight: "Bolder",
        size: "Large",
        color: "Attention",
      },
      {
        type: "FactSet",
        facts: [
          { title: "Employee",    value: data.employeeName },
          { title: "Type",        value: getTypeLabel(data.requestType) },
          { title: "Date",        value: data.displayDate },
          { title: "Rejected By", value: rejectedBy },
          { title: "Time",        value: new Date().toLocaleTimeString("en-IN") },
        ],
      },
    ],
  };
}

// ── Status Card ────────────────────────────────────────────────────────────

export function buildStatusCardContent(
  requestType: string,
  displayDate: string,
  status: "Approved" | "Rejected",
  byName: string
): IAdaptiveCard {
  const isApproved = status === "Approved";
  return {
    $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
    type: "AdaptiveCard",
    version: "1.4",
    body: [
      {
        type: "TextBlock",
        text: isApproved ? "Your request was Approved" : "Your request was Rejected",
        weight: "Bolder",
        size: "Large",
        color: isApproved ? "Good" : "Attention",
      },
      {
        type: "FactSet",
        facts: [
          { title: "Type",     value: getTypeLabel(requestType) },
          { title: "Date",     value: displayDate },
          { title: "Decision", value: `${status} by ${byName}` },
        ],
      },
      {
        type: "TextBlock",
        text: isApproved
          ? "Your leave is confirmed."
          : "Please speak with your manager if you have questions.",
        wrap: true,
        size: "Small",
        color: isApproved ? "Good" : "Warning",
      },
    ],
  };
}

// ── Announcement Card ──────────────────────────────────────────────────────

export function buildAnnouncementCard(record: LeaveRecord): CardActivity {
  return wrap({
    $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
    type: "AdaptiveCard",
    version: "1.4",
    body: [
      {
        type: "TextBlock",
        text: "Workforce Availability Update",
        weight: "Bolder",
        size: "Large",
        color: "Accent",
      },
      {
        type: "FactSet",
        facts: [
          { title: "Employee", value: record.employee },
          { title: "Type",     value: getTypeLabel(record.type) },
          { title: "Date",     value: formatDisplayDate(record.date) },
          { title: "Duration", value: record.duration === "half_day" ? "Half Day" : "Full Day" },
          { title: "Approved By", value: record.approved_by ?? "Manager" },
        ],
      },
    ],
  });
}

// ── Daily Summary Card ─────────────────────────────────────────────────────

export function buildDailySummaryCard(records: LeaveRecord[]): CardActivity {
  const today = new Date().toLocaleDateString("en-IN", {
    weekday: "long", year: "numeric", month: "long", day: "numeric",
  });

  const wfh   = records.filter((r) => r.type === "WFH");
  const leave = records.filter((r) => r.type === "LEAVE");
  const sick  = records.filter((r) => r.type === "SICK");
  const toList = (arr: LeaveRecord[]) =>
    arr.length > 0 ? arr.map((r) => `- ${r.employee}`).join("\n") : "None";

  return wrap({
    $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
    type: "AdaptiveCard",
    version: "1.4",
    body: [
      { type: "TextBlock", text: "Daily Workforce Summary", weight: "Bolder", size: "Large", color: "Accent" },
      { type: "TextBlock", text: today, size: "Small", isSubtle: true },
      {
        type: "ColumnSet",
        columns: [
          {
            type: "Column", width: "stretch",
            items: [
              { type: "TextBlock", text: `WFH (${wfh.length})`, weight: "Bolder" },
              { type: "TextBlock", text: toList(wfh), wrap: true, size: "Small" },
            ],
          },
          {
            type: "Column", width: "stretch",
            items: [
              { type: "TextBlock", text: `On Leave (${leave.length})`, weight: "Bolder" },
              { type: "TextBlock", text: toList(leave), wrap: true, size: "Small" },
            ],
          },
          {
            type: "Column", width: "stretch",
            items: [
              { type: "TextBlock", text: `Sick (${sick.length})`, weight: "Bolder" },
              { type: "TextBlock", text: toList(sick), wrap: true, size: "Small" },
            ],
          },
        ],
      },
      {
        type: "TextBlock",
        text: `Total absent today: ${records.length}`,
        size: "Small",
        isSubtle: true,
        color: records.length > 0 ? "Warning" : "Good",
      },
    ],
  });
}

// ── Help Card ──────────────────────────────────────────────────────────────

export function buildHelpCard(): CardActivity {
  return wrap({
    $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
    type: "AdaptiveCard",
    version: "1.4",
    body: [
      { type: "TextBlock", text: "LeaveAgent - Commands", weight: "Bolder", size: "Large", color: "Accent" },
      {
        type: "FactSet",
        facts: [
          { title: "WFH tomorrow",  value: "Work from home request" },
          { title: "Sick today",    value: "Sick leave" },
          { title: "Leave Friday",  value: "Planned leave" },
          { title: "my requests",   value: "View your last 5 requests" },
          { title: "summary",       value: "Today's workforce availability" },
          { title: "help",          value: "Show this menu" },
        ],
      },
    ],
  });
}

// ── My Requests Card ───────────────────────────────────────────────────────

export function buildMyRequestsCard(userName: string, records: LeaveRecord[]): CardActivity {
  if (records.length === 0) {
    return wrap({
      $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
      type: "AdaptiveCard",
      version: "1.4",
      body: [{ type: "TextBlock", text: "No leave requests found.", wrap: true }],
    });
  }

  return wrap({
    $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
    type: "AdaptiveCard",
    version: "1.4",
    body: [
      { type: "TextBlock", text: `${userName} - Recent Requests`, weight: "Bolder", size: "Large", color: "Accent" },
      {
        type: "FactSet",
        facts: records.map((r) => ({
          title: formatDisplayDate(r.date),
          value: `${getTypeLabel(r.type)} - ${r.status}`,
        })),
      },
    ],
  });
}

// ── Already Processed Card ─────────────────────────────────────────────────

export function buildAlreadyProcessedCardContent(): IAdaptiveCard {
  return {
    $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
    type: "AdaptiveCard",
    version: "1.4",
    body: [
      {
        type: "TextBlock",
        text: "This request has already been processed.",
        color: "Warning",
        wrap: true,
      },
    ],
  };
}