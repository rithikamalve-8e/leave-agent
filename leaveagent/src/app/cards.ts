import { IAdaptiveCard } from "@microsoft/teams.cards";

// ── Shared Types ───────────────────────────────────────────────────────────

export interface ApprovalCardData {
  employeeName:   string;
  employeeEmail:  string;
  requestType:    string;
  date:           string;
  displayDate:    string;
  duration:       string;
  endDate?:       string | null;
  daysCount?:     number;
  lopDays?:       number;
  reason?:        string;
  balanceResult?: any;
}

export interface SimpleCardData {
  employeeName: string;
  requestType:  string;
  date:         string;
  displayDate:  string;
}

export interface CardActivity {
  type: "message";
  attachments: Array<{
    contentType: "application/vnd.microsoft.card.adaptive";
    content:     IAdaptiveCard;
  }>;
}

// Generic leave record shape — works with both excelManager and postgresManager
export interface LeaveRecord {
  employee:          string;
  email?:            string;
  type:              string;
  date:              string;
  end_date?:         string | null;
  duration:          string;
  days_count:        number;
  lop_days?:         number;
  reason?:           string;
  rejection_reason?: string;
  status:            string;
  approved_by?:      string;
  requested_at?:     string | Date;
  updated_at?:       string | Date;
}

// ── Helpers ────────────────────────────────────────────────────────────────

export function formatDisplayDate(isoDate: string): string {
  if (!isoDate) return "Unknown date";
  try {
    const d = new Date(isoDate + "T00:00:00");
    return d.toLocaleDateString("en-IN", {
      weekday: "long", year: "numeric", month: "long", day: "numeric",
    });
  } catch { return isoDate; }
}

export function getTypeLabel(type: string): string {
  switch (type?.toUpperCase()) {
    case "WFH":       return "Work From Home";
    case "LEAVE":     return "Planned Leave";
    case "SICK":      return "Sick Leave";
    case "MATERNITY": return "Maternity Leave";
    case "PATERNITY": return "Paternity Leave";
    case "MARRIAGE":  return "Marriage Leave";
    case "ADOPTION":  return "Adoption Leave";
    default:          return "Leave";
  }
}

export function getTypeEmoji(type: string): string {
  switch (type?.toUpperCase()) {
    case "WFH":       return "🏠";
    case "LEAVE":     return "🌴";
    case "SICK":      return "🤒";
    case "MATERNITY": return "👶";
    case "PATERNITY": return "👨‍👶";
    case "MARRIAGE":  return "💍";
    case "ADOPTION":  return "🏠";
    default:          return "📋";
  }
}

function wrap(content: object): CardActivity {
  return {
    type: "message",
    attachments: [{
      contentType: "application/vnd.microsoft.card.adaptive",
      content: content as IAdaptiveCard
    }],
  };
}

const SCHEMA = "http://adaptivecards.io/schemas/adaptive-card.json";
const VER    = "1.4";

// ── Employee Cards ─────────────────────────────────────────────────────────

export function buildConfirmationCard(
  employeeName:   string,
  requestType:    string,
  displayDate:    string,
  duration:       string,
  endDate?:       string | null,
  daysCount?:     number,
  reason?:        string,
  balanceResult?: any
): CardActivity {
  const facts: any[] = [
    { title: "Employee",     value: employeeName },
    { title: "Type",         value: `${getTypeEmoji(requestType)} ${getTypeLabel(requestType)}` },
    { title: "Date",         value: displayDate },
  ];
  if (endDate)   facts.push({ title: "End Date",     value: endDate });
  if (daysCount) facts.push({ title: "Working Days", value: `${daysCount} day(s)` });
  facts.push(    { title: "Duration",    value: duration });
  if (reason)    facts.push({ title: "Reason",       value: reason });
  if (balanceResult?.hasLop) {
    facts.push({ title: "Leave Balance", value: `${balanceResult.balance} day(s)` });
    facts.push({ title: "Granted",       value: `${balanceResult.granted} day(s)` });
    facts.push({ title: "Loss of Pay",   value: `${balanceResult.lop} day(s) — contact HR` });
  }
  facts.push({ title: "Status", value: "⏳ Awaiting approval" });

  return wrap({
    $schema: SCHEMA, type: "AdaptiveCard", version: VER,
    body: [
      { type: "TextBlock", text: "✅ Request Submitted", weight: "Bolder", size: "Large", color: "Accent" },
      { type: "FactSet", facts },
      { type: "TextBlock", text: "Your approver has been notified and will review shortly.", wrap: true, color: "Good", size: "Small" },
    ],
  });
}

export function buildStatusCardContent(
  requestType:      string,
  displayDate:      string,
  status:           "Approved" | "Rejected",
  byName:           string,
  rejectionReason?: string
): IAdaptiveCard {
  const isApproved = status === "Approved";
  const facts: any[] = [
    { title: "Type",     value: `${getTypeEmoji(requestType)} ${getTypeLabel(requestType)}` },
    { title: "Date",     value: displayDate },
    { title: "Decision", value: `${status} by ${byName}` },
  ];
  if (!isApproved && rejectionReason) {
    facts.push({ title: "Reason", value: rejectionReason });
  }
  return {
    $schema: SCHEMA, type: "AdaptiveCard", version: VER,
    body: [
      {
        type: "TextBlock",
        text: isApproved ? "✅ Your request was Approved" : "❌ Your request was Rejected",
        weight: "Bolder", size: "Large",
        color: isApproved ? "Good" : "Attention",
      },
      { type: "FactSet", facts },
      {
        type: "TextBlock",
        text: isApproved
          ? "Your leave is confirmed. Have a good time!"
          : rejectionReason
            ? `Reason: ${rejectionReason}`
            : "Please speak with your manager if you have questions.",
        wrap: true, size: "Small",
        color: isApproved ? "Good" : "Warning",
      },
    ],
  };
}

export function buildMyRequestsCard(userName: string, records: LeaveRecord[]): CardActivity {
  if (records.length === 0) {
    return wrap({
      $schema: SCHEMA, type: "AdaptiveCard", version: VER,
      body: [
        { type: "TextBlock", text: "📋 My Leave Requests", weight: "Bolder", size: "Large", color: "Accent" },
        { type: "TextBlock", text: "No leave requests found. Submit one by saying 'WFH tomorrow' or 'Sick today'.", wrap: true, color: "Good" },
      ],
    });
  }

  const statusColor = (s: string) =>
    s === "Approved" ? "Good" : s === "Rejected" ? "Attention" : "Warning";

  return wrap({
    $schema: SCHEMA, type: "AdaptiveCard", version: VER,
    body: [
      { type: "TextBlock", text: `📋 ${userName} — Leave Requests`, weight: "Bolder", size: "Large", color: "Accent" },
      ...records.map((r) => ({
        type: "ColumnSet",
        separator: true,
        columns: [
          {
            type: "Column", width: "stretch",
            items: [
              { type: "TextBlock", text: `${getTypeEmoji(r.type)} ${getTypeLabel(r.type)}`, weight: "Bolder", size: "Small" },
              { type: "TextBlock", text: formatDisplayDate(r.date), size: "Small", isSubtle: true, wrap: true },
              ...(r.end_date ? [{ type: "TextBlock", text: `→ ${formatDisplayDate(r.end_date)}`, size: "Small", isSubtle: true }] : []),
            ],
          },
          {
            type: "Column", width: "auto",
            items: [
              { type: "TextBlock", text: r.status, weight: "Bolder", size: "Small", color: statusColor(r.status) },
              { type: "TextBlock", text: `${r.days_count} day(s)`, size: "Small", isSubtle: true },
            ],
          },
        ],
      })),
    ],
  });
}

export function buildLeaveBalanceCard(
  userName:      string,
  balance:       number,
  pendingDays:   number,
  carryForward?: number
): CardActivity {
  const available = Math.max(0, balance - pendingDays);
  const facts: any[] = [
    { title: "Current Balance",   value: `${balance.toFixed(1)} day(s)` },
    { title: "Pending Approvals", value: `${pendingDays.toFixed(1)} day(s)` },
    { title: "Available to Book", value: `${available.toFixed(1)} day(s)` },
  ];
  if (carryForward !== undefined) {
    facts.push({ title: "Carry Forward", value: `${carryForward.toFixed(1)} day(s)` });
  }
  const color = available > 10 ? "Good" : available > 5 ? "Warning" : "Attention";

  return wrap({
    $schema: SCHEMA, type: "AdaptiveCard", version: VER,
    body: [
      { type: "TextBlock", text: `🏖️ Leave Balance — ${userName}`, weight: "Bolder", size: "Large", color: "Accent" },
      { type: "FactSet", facts },
      { type: "TextBlock", text: "💡 WFH does not consume leave balance. LEAVE and SICK are deducted from your annual balance.", wrap: true, size: "Small", isSubtle: true },
      { type: "TextBlock", text: available <= 3 ? "⚠️ Low balance — contact HR for assistance." : "✅ Balance looks good.", wrap: true, size: "Small", color },
    ],
  });
}

export function buildMyStatusCard(record: LeaveRecord): CardActivity {
  const facts: any[] = [
    { title: "Type",     value: `${getTypeEmoji(record.type)} ${getTypeLabel(record.type)}` },
    { title: "Date",     value: formatDisplayDate(record.date) },
  ];
  if (record.end_date) facts.push({ title: "End Date", value: formatDisplayDate(record.end_date) });
  facts.push({ title: "Duration",  value: record.duration === "half_day" ? "Half Day" : record.duration === "multi_day" ? "Multiple Days" : "Full Day" });
  facts.push({ title: "Days",      value: `${record.days_count} day(s)` });
  facts.push({ title: "Status",    value: record.status });
  if (record.approved_by)      facts.push({ title: "Actioned By",      value: record.approved_by });
  if (record.rejection_reason) facts.push({ title: "Rejection Reason", value: record.rejection_reason });

  const color = record.status === "Approved" ? "Good" : record.status === "Rejected" ? "Attention" : "Warning";

  return wrap({
    $schema: SCHEMA, type: "AdaptiveCard", version: VER,
    body: [
      { type: "TextBlock", text: "📋 Request Status", weight: "Bolder", size: "Large", color: "Accent" },
      { type: "FactSet", facts },
      { type: "TextBlock", text: `Status: ${record.status}`, weight: "Bolder", color, size: "Small" },
    ],
  });
}

export function buildHolidaysCard(
  holidays: { date: string; name: string }[],
  month?:   string
): CardActivity {
  const title = month ? `🎉 Holidays — ${month}` : "🎉 Upcoming Holidays";

  if (holidays.length === 0) {
    return wrap({
      $schema: SCHEMA, type: "AdaptiveCard", version: VER,
      body: [
        { type: "TextBlock", text: title, weight: "Bolder", size: "Large", color: "Accent" },
        { type: "TextBlock", text: "No upcoming holidays found.", wrap: true, isSubtle: true },
      ],
    });
  }

  return wrap({
    $schema: SCHEMA, type: "AdaptiveCard", version: VER,
    body: [
      { type: "TextBlock", text: title, weight: "Bolder", size: "Large", color: "Accent" },
      { type: "FactSet", facts: holidays.map((h) => ({ title: formatDisplayDate(h.date), value: h.name })) },
    ],
  });
}

export function buildHelpCard(role: "employee" | "approver" | "hr" = "employee"): CardActivity {
  if (role === "hr")       return buildHRHelpCard();
  if (role === "approver") return buildApproverHelpCard();
  return buildEmployeeHelpCard();
}

function buildEmployeeHelpCard(): CardActivity {
  return wrap({
    $schema: SCHEMA, type: "AdaptiveCard", version: VER,
    body: [
      { type: "TextBlock", text: "🤖 LeaveAgent — Employee Commands", weight: "Bolder", size: "Large", color: "Accent" },
      { type: "TextBlock", text: "Leave Requests", weight: "Bolder", size: "Medium", separator: true },
      { type: "FactSet", facts: [
        { title: "WFH tomorrow",           value: "Work from home request" },
        { title: "Sick today",             value: "Sick leave" },
        { title: "Leave on Friday",        value: "Planned leave" },
        { title: "Leave from 20th to 25th",value: "Multi-day leave" },
      ]},
      { type: "TextBlock", text: "My Info", weight: "Bolder", size: "Medium", separator: true },
      { type: "FactSet", facts: [
        { title: "my requests",      value: "Last 5 requests" },
        { title: "my requests all",  value: "Full request history" },
        { title: "my balance",       value: "Leave balance" },
        { title: "my status [date]", value: "Status of a specific request" },
      ]},
      { type: "TextBlock", text: "General", weight: "Bolder", size: "Medium", separator: true },
      { type: "FactSet", facts: [
        { title: "summary",               value: "Today's workforce availability" },
        { title: "holidays",              value: "Upcoming holidays" },
        { title: "holidays [month]",      value: "Holidays for a specific month" },
        { title: "who is on leave today", value: "See who is out today" },
        { title: "who is wfh today",      value: "See who is WFH today" },
        { title: "delete request [date]", value: "Delete your pending request" },
        { title: "edit request [date]",   value: "Edit your pending request" },
      ]},
    ],
  });
}

function buildApproverHelpCard(): CardActivity {
  return wrap({
    $schema: SCHEMA, type: "AdaptiveCard", version: VER,
    body: [
      { type: "TextBlock", text: "🤖 LeaveAgent — Approver Commands", weight: "Bolder", size: "Large", color: "Accent" },
      { type: "TextBlock", text: "Team Management", weight: "Bolder", size: "Medium", separator: true },
      { type: "FactSet", facts: [
        { title: "team summary",               value: "Team availability today" },
        { title: "team summary [date]",        value: "Team availability on a date" },
        { title: "team requests",              value: "All team requests" },
        { title: "team requests [month]",      value: "Team requests for a month" },
        { title: "pending approvals",          value: "Requests awaiting your action" },
        { title: "balance [name]",             value: "Check a reportee's balance" },
        { title: "leave history [name]",       value: "Reportee's leave history" },
      ]},
      { type: "TextBlock", text: "Queries", weight: "Bolder", size: "Medium", separator: true },
      { type: "FactSet", facts: [
        { title: "who is on leave [date]",  value: "Team leave on a date/range" },
        { title: "who is wfh [date]",       value: "Team WFH on a date/range" },
        { title: "who is available [date]", value: "Team in office on a date" },
      ]},
      { type: "TextBlock", text: "Actions", weight: "Bolder", size: "Medium", separator: true },
      { type: "FactSet", facts: [
        { title: "approve leave [name] [date]",         value: "Approve via command" },
        { title: "reject leave [name] [date] [reason]", value: "Reject via command" },
        { title: "holidays",                            value: "Upcoming holidays" },
        { title: "summary",                             value: "Today's availability" },
      ]},
      { type: "TextBlock", text: "You can also submit your own leave requests normally.", wrap: true, size: "Small", isSubtle: true, separator: true },
    ],
  });
}

function buildHRHelpCard(): CardActivity {
  return wrap({
    $schema: SCHEMA, type: "AdaptiveCard", version: VER,
    body: [
      { type: "TextBlock", text: "🤖 LeaveAgent — HR Commands", weight: "Bolder", size: "Large", color: "Accent" },
      { type: "TextBlock", text: "Queries", weight: "Bolder", size: "Medium", separator: true },
      { type: "FactSet", facts: [
        { title: "all requests",               value: "All leave requests org-wide" },
        { title: "all requests [month]",       value: "Requests for a month" },
        { title: "all requests [name]",        value: "All requests for an employee" },
        { title: "pending",                    value: "All pending requests" },
        { title: "pending [name]",             value: "Pending for an employee" },
        { title: "unactioned",                 value: "Unactioned requests this month" },
        { title: "org summary",                value: "Full org availability today" },
        { title: "org summary [date]",         value: "Org availability on a date" },
        { title: "who is on leave [date]",     value: "Anyone on leave in date range" },
        { title: "who is wfh [date]",          value: "Anyone WFH in date range" },
        { title: "who is available [date]",    value: "Everyone in office on a date" },
        { title: "leave history [name]",       value: "Full history for any employee" },
        { title: "balance [name]",             value: "Any employee's balance" },
        { title: "view employee [name]",       value: "Employee profile" },
        { title: "unregistered",               value: "Employees not yet on bot" },
        { title: "team [approver name]",       value: "View an approver's reportees" },
        { title: "org chart",                  value: "Full org structure" },
        { title: "audit log",                  value: "Last 50 HR actions" },
        { title: "audit log [date] to [date]", value: "Audit log for a date range" },
      ]},
      { type: "TextBlock", text: "Leave Management", weight: "Bolder", size: "Medium", separator: true },
      { type: "FactSet", facts: [
        { title: "add leave for [name] [type] [date]",  value: "Submit on behalf (auto-approved)" },
        { title: "approve leave [name] [date]",         value: "Manually approve any request" },
        { title: "reject leave [name] [date] [reason]", value: "Manually reject any request" },
        { title: "delete request [name] [date]",        value: "Delete any request" },
        { title: "restore request [name] [date]",       value: "Restore a deleted request" },
        { title: "approve unactioned [name]",           value: "Approve unactioned for one employee" },
        { title: "approve all unactioned",              value: "Bulk approve all unactioned" },
        { title: "reject all unactioned [reason]",      value: "Bulk reject all unactioned" },
        { title: "remind approvers",                    value: "Manually trigger approver reminders" },
      ]},
      { type: "TextBlock", text: "Employee Management", weight: "Bolder", size: "Medium", separator: true },
      { type: "FactSet", facts: [
        { title: "adjust balance [name] [+/-days] [reason]", value: "Add or deduct leave days" },
        { title: "set balance [name] [days]",                value: "Set exact leave balance" },
        { title: "reset balances [year]",                    value: "Reset all balances to 22" },
        { title: "add employee [details]",                   value: "Add new employee" },
        { title: "update employee [name] [field] [value]",   value: "Update employee field" },
        { title: "deactivate employee [name]",               value: "Mark employee inactive" },
      ]},
      { type: "TextBlock", text: "Holidays", weight: "Bolder", size: "Medium", separator: true },
      { type: "FactSet", facts: [
        { title: "add holiday [date] [name]",           value: "Add a company holiday" },
        { title: "edit holiday [date] [new name]",      value: "Edit holiday name" },
        { title: "reschedule holiday [name] to [date]", value: "Move holiday to new date" },
        { title: "delete holiday [date]",               value: "Remove a holiday" },
        { title: "holidays",                            value: "List upcoming holidays" },
      ]},
      { type: "TextBlock", text: "Reports", weight: "Bolder", size: "Medium", separator: true },
      { type: "FactSet", facts: [
        { title: "download report [month] [year]", value: "Monthly attendance report (xlsx)" },
        { title: "download report ytd",            value: "Year-to-date report" },
      ]},
    ],
  });
}

// ── Approver Cards ─────────────────────────────────────────────────────────

export function buildApprovalCardContent(data: ApprovalCardData): IAdaptiveCard {
  const facts: any[] = [
    { title: "Employee",     value: data.employeeName },
    { title: "Email",        value: data.employeeEmail },
    { title: "Type",         value: `${getTypeEmoji(data.requestType)} ${getTypeLabel(data.requestType)}` },
    { title: "Date",         value: data.displayDate },
    ...(data.endDate   ? [{ title: "End Date",     value: data.endDate }]              : []),
    ...(data.daysCount ? [{ title: "Working Days", value: `${data.daysCount} day(s)` }] : []),
    { title: "Duration",     value: data.duration },
    ...(data.reason    ? [{ title: "Reason",       value: data.reason }]               : []),
    ...(data.balanceResult?.hasLop ? [
      { title: "Leave Balance", value: `${data.balanceResult.balance} day(s)` },
      { title: "Granted",       value: `${data.balanceResult.granted} day(s)` },
      { title: "LOP",           value: `${data.balanceResult.lop} day(s)` },
    ] : []),
    ...(data.lopDays && data.lopDays > 0 ? [{ title: "Loss of Pay", value: `${data.lopDays} day(s)` }] : []),
  ];

  return {
    $schema: SCHEMA, type: "AdaptiveCard", version: VER,
    body: [
      { type: "TextBlock", text: "⏳ Leave Approval Required", weight: "Bolder", size: "Large", color: "Warning" },
      { type: "FactSet", facts },
      { type: "TextBlock", text: "Please approve or reject this request.", wrap: true, size: "Small" },
    ],
    actions: [
      {
        type: "Action.Submit", title: "✅ Approve", style: "positive",
        data: { action: "approve", employeeName: data.employeeName, date: data.date, requestType: data.requestType },
      },
      {
        type: "Action.Submit", title: "❌ Reject", style: "destructive",
        data: { action: "reject", employeeName: data.employeeName, date: data.date, requestType: data.requestType },
      },
    ],
  };
}

export function buildApprovedCardContent(data: SimpleCardData, approvedBy: string): IAdaptiveCard {
  return {
    $schema: SCHEMA, type: "AdaptiveCard", version: VER,
    body: [
      { type: "TextBlock", text: "✅ Request Approved", weight: "Bolder", size: "Large", color: "Good" },
      { type: "FactSet", facts: [
        { title: "Employee",    value: data.employeeName },
        { title: "Type",        value: `${getTypeEmoji(data.requestType)} ${getTypeLabel(data.requestType)}` },
        { title: "Date",        value: data.displayDate },
        { title: "Approved By", value: approvedBy },
        { title: "Time",        value: new Date().toLocaleTimeString("en-IN") },
      ]},
    ],
  };
}

export function buildRejectedCardContent(
  data:             SimpleCardData,
  rejectedBy:       string,
  rejectionReason?: string
): IAdaptiveCard {
  const facts: any[] = [
    { title: "Employee",    value: data.employeeName },
    { title: "Type",        value: `${getTypeEmoji(data.requestType)} ${getTypeLabel(data.requestType)}` },
    { title: "Date",        value: data.displayDate },
    { title: "Rejected By", value: rejectedBy },
    { title: "Time",        value: new Date().toLocaleTimeString("en-IN") },
  ];
  if (rejectionReason) facts.push({ title: "Reason", value: rejectionReason });

  return {
    $schema: SCHEMA, type: "AdaptiveCard", version: VER,
    body: [
      { type: "TextBlock", text: "❌ Request Rejected", weight: "Bolder", size: "Large", color: "Attention" },
      { type: "FactSet", facts },
    ],
  };
}

export function buildRejectionReasonPromptCard(
  employeeName: string,
  date:         string,
  requestType:  string,
  displayDate:  string
): CardActivity {
  return wrap({
    $schema: SCHEMA, type: "AdaptiveCard", version: VER,
    body: [
      { type: "TextBlock", text: "❌ Rejecting Request", weight: "Bolder", size: "Large", color: "Attention" },
      { type: "FactSet", facts: [
        { title: "Employee", value: employeeName },
        { title: "Type",     value: getTypeLabel(requestType) },
        { title: "Date",     value: displayDate },
      ]},
      { type: "TextBlock", text: "Please provide a reason for rejection:", wrap: true },
      {
        type: "Input.Text", id: "rejectionReason",
        placeholder: "e.g. Project deadline conflict, insufficient notice...",
        isMultiline: true, maxLength: 300,
      },
    ],
    actions: [
      {
        type: "Action.Submit", title: "Confirm Rejection", style: "destructive",
        data: { action: "confirm_reject", employeeName, date, requestType },
      },
      {
        type: "Action.Submit", title: "Cancel",
        data: { action: "cancel_reject", employeeName, date, requestType },
      },
    ],
  });
}

export function buildTeamRequestsCard(
  approverName: string,
  records:      LeaveRecord[],
  title?:       string
): CardActivity {
  const heading = title ?? `👥 Team Requests — ${approverName}`;

  if (records.length === 0) {
    return wrap({
      $schema: SCHEMA, type: "AdaptiveCard", version: VER,
      body: [
        { type: "TextBlock", text: heading, weight: "Bolder", size: "Large", color: "Accent" },
        { type: "TextBlock", text: "No requests found for your team.", wrap: true, isSubtle: true },
      ],
    });
  }

  const statusColor = (s: string) =>
    s === "Approved" ? "Good" : s === "Rejected" ? "Attention" : "Warning";

  return wrap({
    $schema: SCHEMA, type: "AdaptiveCard", version: VER,
    body: [
      { type: "TextBlock", text: heading, weight: "Bolder", size: "Large", color: "Accent" },
      { type: "TextBlock", text: `${records.length} request(s)`, size: "Small", isSubtle: true },
      ...records.map((r) => ({
        type: "ColumnSet", separator: true,
        columns: [
          {
            type: "Column", width: "stretch",
            items: [
              { type: "TextBlock", text: `${getTypeEmoji(r.type)} ${r.employee}`, weight: "Bolder", size: "Small" },
              { type: "TextBlock", text: `${getTypeLabel(r.type)} · ${formatDisplayDate(r.date)}`, size: "Small", isSubtle: true, wrap: true },
            ],
          },
          {
            type: "Column", width: "auto",
            items: [
              { type: "TextBlock", text: r.status, weight: "Bolder", size: "Small", color: statusColor(r.status) },
              { type: "TextBlock", text: `${r.days_count} day(s)`, size: "Small", isSubtle: true },
            ],
          },
        ],
      })),
    ],
  });
}

export function buildPendingApprovalsCard(
  approverName: string,
  records:      LeaveRecord[]
): CardActivity {
  if (records.length === 0) {
    return wrap({
      $schema: SCHEMA, type: "AdaptiveCard", version: VER,
      body: [
        { type: "TextBlock", text: "⏳ Pending Approvals", weight: "Bolder", size: "Large", color: "Accent" },
        { type: "TextBlock", text: "No pending requests. You're all caught up! 🎉", wrap: true, color: "Good" },
      ],
    });
  }

  return wrap({
    $schema: SCHEMA, type: "AdaptiveCard", version: VER,
    body: [
      { type: "TextBlock", text: `⏳ Pending Approvals — ${approverName}`, weight: "Bolder", size: "Large", color: "Warning" },
      { type: "TextBlock", text: `${records.length} request(s) awaiting your action`, size: "Small", color: "Warning" },
      ...records.map((r) => ({
        type: "ColumnSet", separator: true,
        columns: [{
          type: "Column", width: "stretch",
          items: [
            { type: "TextBlock", text: `${getTypeEmoji(r.type)} ${r.employee}`, weight: "Bolder", size: "Small" },
            { type: "TextBlock", text: `${getTypeLabel(r.type)} · ${formatDisplayDate(r.date)}`, size: "Small", isSubtle: true, wrap: true },
            { type: "TextBlock", text: `${r.days_count} day(s)`, size: "Small", isSubtle: true },
          ],
        }],
      })),
      { type: "TextBlock", text: "Use the approval cards sent to you, or type: approve leave [name] [date]", wrap: true, size: "Small", isSubtle: true, separator: true },
    ],
  });
}

export function buildWhoIsOnLeaveCard(
  records:   LeaveRecord[],
  dateLabel: string,
  type?:     "leave" | "wfh" | "all"
): CardActivity {
  const typeTitle =
    type === "wfh"   ? "🏠 WFH" :
    type === "leave" ? "🌴 On Leave" : "📋 Absent";

  if (records.length === 0) {
    return wrap({
      $schema: SCHEMA, type: "AdaptiveCard", version: VER,
      body: [
        { type: "TextBlock", text: `${typeTitle} — ${dateLabel}`, weight: "Bolder", size: "Large", color: "Accent" },
        { type: "TextBlock", text: "Everyone is available! 🎉", wrap: true, color: "Good" },
      ],
    });
  }

  return wrap({
    $schema: SCHEMA, type: "AdaptiveCard", version: VER,
    body: [
      { type: "TextBlock", text: `${typeTitle} — ${dateLabel}`, weight: "Bolder", size: "Large", color: "Accent" },
      { type: "TextBlock", text: `${records.length} person(s)`, size: "Small", isSubtle: true },
      { type: "FactSet", facts: records.map((r) => ({
        title: r.employee,
        value: `${getTypeEmoji(r.type)} ${getTypeLabel(r.type)} · ${r.days_count} day(s)`,
      }))},
    ],
  });
}

export function buildApproverReminderCard(
  approverName: string,
  records:      LeaveRecord[],
  month:        string
): CardActivity {
  return wrap({
    $schema: SCHEMA, type: "AdaptiveCard", version: VER,
    body: [
      { type: "TextBlock", text: "⚠️ Month-End Reminder", weight: "Bolder", size: "Large", color: "Warning" },
      { type: "TextBlock", text: `Hi ${approverName}, the following requests from ${month} are still pending:`, wrap: true },
      { type: "FactSet", facts: records.map((r) => ({
        title: r.employee,
        value: `${getTypeLabel(r.type)} on ${formatDisplayDate(r.date)}`,
      }))},
      { type: "TextBlock", text: "⏰ Please action these by end of day — unactioned requests will be escalated to HR.", wrap: true, color: "Attention", size: "Small" },
    ],
  });
}

// ── HR Cards ───────────────────────────────────────────────────────────────

export function buildHRAlertCard(
  event:        "submitted" | "approved" | "rejected",
  employeeName: string,
  requestType:  string,
  displayDate:  string,
  actionBy?:    string,
  reason?:      string
): CardActivity {
  const titles = { submitted: "📬 New Leave Request", approved: "✅ Request Approved", rejected: "❌ Request Rejected" };
  const colors = { submitted: "Accent", approved: "Good", rejected: "Attention" };

  const facts: any[] = [
    { title: "Employee", value: employeeName },
    { title: "Type",     value: `${getTypeEmoji(requestType)} ${getTypeLabel(requestType)}` },
    { title: "Date",     value: displayDate },
  ];
  if (actionBy) facts.push({ title: event === "submitted" ? "Requested By" : "Actioned By", value: actionBy });
  if (reason)   facts.push({ title: "Reason", value: reason });

  return wrap({
    $schema: SCHEMA, type: "AdaptiveCard", version: VER,
    body: [
      { type: "TextBlock", text: titles[event], weight: "Bolder", size: "Large", color: colors[event] as any },
      { type: "FactSet", facts },
    ],
  });
}

export function buildHRTakeoverCard(records: LeaveRecord[], month: string): CardActivity {
  return wrap({
    $schema: SCHEMA, type: "AdaptiveCard", version: VER,
    body: [
      { type: "TextBlock", text: "🚨 Unactioned Requests — HR Escalation", weight: "Bolder", size: "Large", color: "Attention" },
      { type: "TextBlock", text: `The following ${month} requests were not actioned by their approvers:`, wrap: true },
      { type: "FactSet", facts: records.map((r) => ({
        title: r.employee,
        value: `${getTypeLabel(r.type)} on ${formatDisplayDate(r.date)}`,
      }))},
      { type: "TextBlock", text: "Use: approve leave [name] [date] OR reject leave [name] [date] [reason]", wrap: true, size: "Small", isSubtle: true, separator: true },
    ],
  });
}

export function buildAllRequestsCard(records: LeaveRecord[], title?: string): CardActivity {
  const heading = title ?? "📋 All Leave Requests";

  if (records.length === 0) {
    return wrap({
      $schema: SCHEMA, type: "AdaptiveCard", version: VER,
      body: [
        { type: "TextBlock", text: heading, weight: "Bolder", size: "Large", color: "Accent" },
        { type: "TextBlock", text: "No requests found.", wrap: true, isSubtle: true },
      ],
    });
  }

  const buildSection = (items: LeaveRecord[], label: string): any[] => {
    if (items.length === 0) return [];
    return [
      { type: "TextBlock", text: `${label} (${items.length})`, weight: "Bolder", size: "Small", separator: true },
      { type: "FactSet", facts: items.map((r) => ({
        title: `${r.employee} · ${formatDisplayDate(r.date)}`,
        value: `${getTypeEmoji(r.type)} ${getTypeLabel(r.type)} · ${r.days_count} day(s)`,
      }))},
    ];
  };

  return wrap({
    $schema: SCHEMA, type: "AdaptiveCard", version: VER,
    body: [
      { type: "TextBlock", text: heading, weight: "Bolder", size: "Large", color: "Accent" },
      { type: "TextBlock", text: `Total: ${records.length}`, size: "Small", isSubtle: true },
      ...buildSection(records.filter((r) => r.status === "Pending"),  "⏳ Pending"),
      ...buildSection(records.filter((r) => r.status === "Approved"), "✅ Approved"),
      ...buildSection(records.filter((r) => r.status === "Rejected"), "❌ Rejected"),
      ...buildSection(records.filter((r) => r.status === "Deleted"),  "🗑️ Deleted"),
    ],
  });
}

export function buildAuditLogCard(
  entries:    { timestamp: Date; hr_name: string; action: string; target_employee: string | null; details: string }[],
  dateRange?: string
): CardActivity {
  const title = dateRange ? `🔍 Audit Log — ${dateRange}` : "🔍 Audit Log — Last 50 Actions";

  if (entries.length === 0) {
    return wrap({
      $schema: SCHEMA, type: "AdaptiveCard", version: VER,
      body: [
        { type: "TextBlock", text: title, weight: "Bolder", size: "Large", color: "Accent" },
        { type: "TextBlock", text: "No audit entries found.", wrap: true, isSubtle: true },
      ],
    });
  }

  const actionEmoji = (a: string) =>
    a === "balance_adjust"   ? "💰" :
    a === "delete_request"   ? "🗑️" :
    a === "add_holiday"      ? "🎉" :
    a === "add_leave_behalf" ? "📝" :
    a === "restore_request"  ? "♻️" : "📋";

  return wrap({
    $schema: SCHEMA, type: "AdaptiveCard", version: VER,
    body: [
      { type: "TextBlock", text: title, weight: "Bolder", size: "Large", color: "Accent" },
      { type: "TextBlock", text: `${entries.length} entries`, size: "Small", isSubtle: true },
      ...entries.slice(0, 20).map((e) => ({
        type: "ColumnSet", separator: true,
        columns: [{
          type: "Column", width: "stretch",
          items: [
            { type: "TextBlock", text: `${actionEmoji(e.action)} ${e.hr_name} — ${e.action.replace(/_/g, " ")}`, weight: "Bolder", size: "Small" },
            { type: "TextBlock", text: e.details, size: "Small", isSubtle: true, wrap: true },
            { type: "TextBlock", text: new Date(e.timestamp).toLocaleString("en-IN"), size: "Small", isSubtle: true },
          ],
        }],
      })),
      ...(entries.length > 20 ? [{ type: "TextBlock", text: `... and ${entries.length - 20} more. Download report for full log.`, wrap: true, size: "Small", isSubtle: true }] : []),
    ],
  });
}

export function buildEmployeeProfileCard(emp: {
  name: string; email: string; role: string; bot_role: string;
  manager?: string | null; teamlead?: string | null;
  leave_balance: number; carry_forward?: number;
  teams_id?: string | null;
}): CardActivity {
  const facts: any[] = [
    { title: "Name",          value: emp.name },
    { title: "Email",         value: emp.email },
    { title: "Org Role",      value: emp.role },
    { title: "Bot Role",      value: emp.bot_role },
    { title: "Leave Balance", value: `${emp.leave_balance} day(s)` },
  ];
  if (emp.carry_forward !== undefined) facts.push({ title: "Carry Forward", value: `${emp.carry_forward} day(s)` });
  if (emp.manager)                     facts.push({ title: "Manager",       value: emp.manager });
  if (emp.teamlead)                    facts.push({ title: "Team Lead",     value: emp.teamlead });
  facts.push({ title: "Teams ID", value: emp.teams_id ? "✅ Registered" : "❌ Not registered" });

  return wrap({
    $schema: SCHEMA, type: "AdaptiveCard", version: VER,
    body: [
      { type: "TextBlock", text: `👤 ${emp.name}`, weight: "Bolder", size: "Large", color: "Accent" },
      { type: "FactSet", facts },
    ],
  });
}

export function buildBalanceAdjustedCard(
  employeeName: string,
  adjustment:   number,
  newBalance:   number,
  reason:       string,
  adjustedBy:   string
): CardActivity {
  return wrap({
    $schema: SCHEMA, type: "AdaptiveCard", version: VER,
    body: [
      { type: "TextBlock", text: "💰 Leave Balance Updated", weight: "Bolder", size: "Large", color: "Accent" },
      { type: "FactSet", facts: [
        { title: "Employee",    value: employeeName },
        { title: "Adjustment",  value: `${adjustment > 0 ? "+" : ""}${adjustment} day(s)` },
        { title: "New Balance", value: `${newBalance} day(s)` },
        { title: "Reason",      value: reason },
        { title: "Updated By",  value: adjustedBy },
        { title: "Time",        value: new Date().toLocaleString("en-IN") },
      ]},
    ],
  });
}

export function buildDeletedNotificationCard(
  employeeName: string,
  requestType:  string,
  displayDate:  string,
  deletedBy:    string,
  reason:       string
): CardActivity {
  return wrap({
    $schema: SCHEMA, type: "AdaptiveCard", version: VER,
    body: [
      { type: "TextBlock", text: "🗑️ Leave Request Deleted", weight: "Bolder", size: "Large", color: "Attention" },
      { type: "FactSet", facts: [
        { title: "Employee",   value: employeeName },
        { title: "Type",       value: getTypeLabel(requestType) },
        { title: "Date",       value: displayDate },
        { title: "Deleted By", value: deletedBy },
        { title: "Reason",     value: reason },
      ]},
      { type: "TextBlock", text: "Contact HR if you believe this was done in error.", wrap: true, size: "Small", isSubtle: true },
    ],
  });
}

export function buildHolidayAnnouncementCard(
  date:    string,
  name:    string,
  addedBy: string,
  action:  "added" | "edited" | "rescheduled" | "deleted" = "added"
): CardActivity {
  const titles = {
    added:       "🎉 New Holiday Added",
    edited:      "✏️ Holiday Updated",
    rescheduled: "📅 Holiday Rescheduled",
    deleted:     "🗑️ Holiday Removed",
  };

  return wrap({
    $schema: SCHEMA, type: "AdaptiveCard", version: VER,
    body: [
      { type: "TextBlock", text: titles[action], weight: "Bolder", size: "Large", color: "Accent" },
      { type: "FactSet", facts: [
        { title: "Holiday",  value: name },
        { title: "Date",     value: formatDisplayDate(date) },
        { title: "Added By", value: addedBy },
      ]},
      { type: "TextBlock", text: "Mark your calendars! 🗓️", wrap: true, size: "Small", color: "Good" },
    ],
  });
}

export function buildUnregisteredCard(
  employees: { name: string; email: string; role: string }[]
): CardActivity {
  if (employees.length === 0) {
    return wrap({
      $schema: SCHEMA, type: "AdaptiveCard", version: VER,
      body: [
        { type: "TextBlock", text: "✅ All Employees Registered", weight: "Bolder", size: "Large", color: "Good" },
        { type: "TextBlock", text: "Everyone has messaged the bot and is registered.", wrap: true },
      ],
    });
  }

  return wrap({
    $schema: SCHEMA, type: "AdaptiveCard", version: VER,
    body: [
      { type: "TextBlock", text: "⚠️ Unregistered Employees", weight: "Bolder", size: "Large", color: "Warning" },
      { type: "TextBlock", text: `${employees.length} employee(s) have not yet messaged the bot.`, wrap: true },
      { type: "FactSet", facts: employees.map((e) => ({ title: e.name, value: `${e.email} · ${e.role}` })) },
      { type: "TextBlock", text: "Ask them to open LeaveAgent in Teams and send any message.", wrap: true, size: "Small", isSubtle: true, separator: true },
    ],
  });
}

export function buildOrgChartCard(
  approvers: { name: string; reportees: string[] }[]
): CardActivity {
  return wrap({
    $schema: SCHEMA, type: "AdaptiveCard", version: VER,
    body: [
      { type: "TextBlock", text: "🏢 Org Chart", weight: "Bolder", size: "Large", color: "Accent" },
      ...approvers.map((a) => ({
        type: "Container", separator: true,
        items: [
          { type: "TextBlock", text: `👤 ${a.name}`, weight: "Bolder", size: "Small" },
          {
            type: "TextBlock",
            text: a.reportees.length > 0
              ? a.reportees.map((r) => `  └ ${r}`).join("\n")
              : "  └ No direct reportees",
            size: "Small", isSubtle: true, wrap: true,
          },
        ],
      })),
    ],
  });
}

export function buildUnactionedCard(records: LeaveRecord[], forHR = false): CardActivity {
  const title = forHR ? "🚨 Unactioned Requests This Month" : "⏳ Your Unactioned Requests";

  if (records.length === 0) {
    return wrap({
      $schema: SCHEMA, type: "AdaptiveCard", version: VER,
      body: [
        { type: "TextBlock", text: title, weight: "Bolder", size: "Large", color: "Good" },
        { type: "TextBlock", text: "No unactioned requests this month. 🎉", wrap: true },
      ],
    });
  }

  return wrap({
    $schema: SCHEMA, type: "AdaptiveCard", version: VER,
    body: [
      { type: "TextBlock", text: title, weight: "Bolder", size: "Large", color: "Warning" },
      { type: "TextBlock", text: `${records.length} request(s) pending action`, size: "Small", color: "Warning" },
      { type: "FactSet", facts: records.map((r) => ({
        title: forHR ? r.employee : formatDisplayDate(r.date),
        value: forHR
          ? `${getTypeLabel(r.type)} on ${formatDisplayDate(r.date)}`
          : `${getTypeLabel(r.type)} · ${r.days_count} day(s)`,
      }))},
      {
        type: "TextBlock",
        text: forHR
          ? "Use: approve all unactioned OR approve unactioned [name]"
          : "These requests are awaiting your approval.",
        wrap: true, size: "Small", isSubtle: true, separator: true,
      },
    ],
  });
}

// ── Shared Cards ───────────────────────────────────────────────────────────

export function buildDailySummaryCard(records: LeaveRecord[]): CardActivity {
  const today = new Date().toLocaleDateString("en-IN", {
    weekday: "long", year: "numeric", month: "long", day: "numeric",
  });

  const wfh   = records.filter((r) => r.type === "WFH");
  const leave = records.filter((r) => r.type === "LEAVE");
  const sick  = records.filter((r) => r.type === "SICK");
  const other = records.filter((r) => !["WFH","LEAVE","SICK"].includes(r.type?.toUpperCase()));

  const toList = (arr: LeaveRecord[]) =>
    arr.length > 0 ? arr.map((r) => `• ${r.employee}`).join("\n") : "None";

  return wrap({
    $schema: SCHEMA, type: "AdaptiveCard", version: VER,
    body: [
      { type: "TextBlock", text: "📋 Daily Workforce Summary", weight: "Bolder", size: "Large", color: "Accent" },
      { type: "TextBlock", text: today, size: "Small", isSubtle: true },
      {
        type: "ColumnSet",
        columns: [
          { type: "Column", width: "stretch", items: [{ type: "TextBlock", text: `🏠 WFH (${wfh.length})`, weight: "Bolder" }, { type: "TextBlock", text: toList(wfh), wrap: true, size: "Small" }] },
          { type: "Column", width: "stretch", items: [{ type: "TextBlock", text: `🌴 Leave (${leave.length})`, weight: "Bolder" }, { type: "TextBlock", text: toList(leave), wrap: true, size: "Small" }] },
          { type: "Column", width: "stretch", items: [{ type: "TextBlock", text: `🤒 Sick (${sick.length})`, weight: "Bolder" }, { type: "TextBlock", text: toList(sick), wrap: true, size: "Small" }] },
        ],
      },
      ...(other.length > 0 ? [{ type: "TextBlock", text: `Other: ${other.map((r) => r.employee).join(", ")}`, size: "Small", isSubtle: true }] : []),
      {
        type: "TextBlock",
        text: records.length === 0 ? "✅ Everyone is available today." : `Total absent: ${records.length}`,
        size: "Small", isSubtle: true,
        color: records.length > 0 ? "Warning" : "Good",
      },
    ],
  });
}

export interface PreviewCardData {
  employeeName: string; requestType: string; date: string;
  displayDate: string; endDate?: string | null;
  daysCount: number; duration: string; reason?: string; balanceResult?: any;
}

export function buildPreviewCard(data: PreviewCardData): CardActivity {
  const facts: any[] = [
    { title: "Employee",     value: data.employeeName },
    { title: "Type",         value: `${getTypeEmoji(data.requestType)} ${getTypeLabel(data.requestType)}` },
    { title: "Date",         value: data.displayDate },
  ];
  if (data.endDate)   facts.push({ title: "End Date",     value: data.endDate });
  if (data.daysCount) facts.push({ title: "Working Days", value: `${data.daysCount} day(s)` });
  facts.push(         { title: "Duration",    value: data.duration });
  if (data.reason)    facts.push({ title: "Reason",       value: data.reason });

  return wrap({
    $schema: SCHEMA, type: "AdaptiveCard", version: VER,
    body: [
      { type: "TextBlock", text: "📋 Review Your Request", weight: "Bolder", size: "Large", color: "Accent" },
      { type: "FactSet", facts },
      { type: "TextBlock", text: "Please confirm or edit before sending to your approver.", wrap: true, size: "Small", isSubtle: true },
    ],
    actions: [
      { type: "Action.Submit", title: "✅ Confirm & Send", style: "positive", data: { action: "preview_confirm" } },
      { type: "Action.Submit", title: "✏️ Edit",                                data: { action: "preview_edit"    } },
      { type: "Action.Submit", title: "❌ Cancel",          style: "destructive", data: { action: "preview_cancel"  } },
    ],
  });
}

export function buildCancelledCard(): CardActivity {
  return wrap({
    $schema: SCHEMA, type: "AdaptiveCard", version: VER,
    body: [
      { type: "TextBlock", text: "❌ Request Cancelled", weight: "Bolder", size: "Large", color: "Attention" },
      { type: "TextBlock", text: "Your request was cancelled. Nothing was submitted.", wrap: true },
    ],
  });
}

export function buildAnnouncementCard(record: {
  employee:      string;
  email?:        string;
  type:          string;
  date:          string;
  end_date?:     string | null;
  duration:      string;
  days_count:    number;
  lop_days?:     number;
  status:        string;
  approved_by?:  string;
  requested_at?: string;
}): CardActivity {
  const facts: any[] = [
    { title: "Employee",    value: record.employee },
    { title: "Type",        value: `${getTypeEmoji(record.type)} ${getTypeLabel(record.type)}` },
    { title: "Date",        value: formatDisplayDate(record.date) },
  ];
  if (record.end_date)    facts.push({ title: "End Date",    value: formatDisplayDate(record.end_date) });
  facts.push({ title: "Duration",    value: record.duration === "half_day" ? "Half Day" : record.duration === "multi_day" ? "Multiple Days" : "Full Day" });
  facts.push({ title: "Days",        value: `${record.days_count} day(s)` });
  if (record.approved_by) facts.push({ title: "Approved By", value: record.approved_by });

  return wrap({
    $schema: SCHEMA, type: "AdaptiveCard", version: VER,
    body: [
      { type: "TextBlock", text: "📅 Workforce Availability Update", weight: "Bolder", size: "Large", color: "Accent" },
      { type: "FactSet", facts },
    ],
  });
}

export function buildAlreadyProcessedCardContent(): IAdaptiveCard {
  return {
    $schema: SCHEMA, type: "AdaptiveCard", version: VER,
    body: [
      { type: "TextBlock", text: "⚠️ Already Processed", weight: "Bolder", color: "Warning" },
      { type: "TextBlock", text: "This request has already been processed.", wrap: true, size: "Small" },
    ],
  };
}

export function buildErrorCard(message: string): CardActivity {
  return wrap({
    $schema: SCHEMA, type: "AdaptiveCard", version: VER,
    body: [
      { type: "TextBlock", text: "❌ Error", weight: "Bolder", size: "Large", color: "Attention" },
      { type: "TextBlock", text: message, wrap: true },
    ],
  });
}

export function buildSuccessCard(title: string, message: string): CardActivity {
  return wrap({
    $schema: SCHEMA, type: "AdaptiveCard", version: VER,
    body: [
      { type: "TextBlock", text: `✅ ${title}`, weight: "Bolder", size: "Large", color: "Good" },
      { type: "TextBlock", text: message, wrap: true },
    ],
  });
}
