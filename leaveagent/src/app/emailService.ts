/**
 * emailService.ts
 * Sends email notifications via Outlook / Office 365 SMTP.
 *
 * Approval chain depends on role (from Employees.xlsx):
 *
 *   role = "employee"  → Team Lead approves
 *                         To: Team Lead | CC: Manager, HR
 *
 *   role = "teamlead"  → Manager approves
 *                         To: Manager  | CC: HR
 *
 * On decision (approved/rejected):
 *   role = "employee"  → To: Employee | CC: Manager, HR
 *   role = "teamlead"  → To: Employee | CC: HR
 *
 * HR is global — set HR_EMAIL in env/.env.local
 */

import nodemailer from "nodemailer";

const transporter = nodemailer.createTransport({
  host:   "smtp.office365.com",
  port:   587,
  secure: false,
  auth: {
    user: process.env.EMAIL_USER,
    pass: process.env.EMAIL_PASS,
  },
  tls: { ciphers: "SSLv3" },
});

export interface LeaveEmailData {
  employeeName:   string;
  employeeEmail:  string;
  role:           "employee" | "teamlead";
  managerName:    string;
  managerEmail:   string;
  teamleadName:   string;
  teamleadEmail:  string;
  requestType:    string;
  displayDate:    string;
  duration:       string;
  status?:        "Pending" | "Approved" | "Rejected";
  decidedBy?:     string;
}

function getTypeLabel(type: string): string {
  switch (type?.toUpperCase()) {
    case "WFH":   return "Work From Home";
    case "LEAVE": return "Planned Leave";
    case "SICK":  return "Sick Leave";
    default:      return "Leave Request";
  }
}

function tableRow(label: string, value: string, shaded: boolean): string {
  const bg = shaded ? "background:#f5f5f5;" : "";
  return `
    <tr>
      <td style="padding:10px;font-weight:bold;width:40%;border:1px solid #ddd;${bg}">${label}</td>
      <td style="padding:10px;border:1px solid #ddd;${bg}">${value}</td>
    </tr>`;
}

function buildTable(rows: Array<[string, string]>): string {
  return `<table style="width:100%;border-collapse:collapse;margin:16px 0;">
    ${rows.map(([l, v], i) => tableRow(l, v, i % 2 === 0)).join("")}
  </table>`;
}

// ── Email 1: New Request ───────────────────────────────────────────────────

export async function sendRequestNotificationEmail(data: LeaveEmailData): Promise<void> {
  const hrEmail = process.env.HR_EMAIL ?? "";

  let toEmail:   string;
  let toName:    string;
  let ccEmails:  string[];

  if (data.role === "employee") {
    // Employee: Team Lead approves, CC Manager + HR
    toEmail  = data.teamleadEmail;
    toName   = data.teamleadName;
    ccEmails = [data.managerEmail, hrEmail].filter(Boolean);
  } else {
    // Team Lead: Manager approves, CC HR only
    toEmail  = data.managerEmail;
    toName   = data.managerName;
    ccEmails = [hrEmail].filter(Boolean);
  }

  const subject = `Leave Request: ${data.employeeName} - ${getTypeLabel(data.requestType)} on ${data.displayDate}`;

  const rows: Array<[string, string]> = [
    ["Employee",   data.employeeName],
    ["Email",      data.employeeEmail],
    ["Role",       data.role === "teamlead" ? "Team Lead" : "Employee"],
    ["Type",       getTypeLabel(data.requestType)],
    ["Date",       data.displayDate],
    ["Duration",   data.duration],
    ["Manager",    data.managerName],
    ...(data.role === "employee"
      ? [["Team Lead", data.teamleadName] as [string, string]]
      : []),
    ["Status",     "Pending Approval"],
  ];

  const html = `
    <div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto;">
      <div style="background:#0078d4;padding:20px;border-radius:8px 8px 0 0;">
        <h2 style="color:white;margin:0;">Leave Request Submitted</h2>
      </div>
      <div style="border:1px solid #e0e0e0;border-top:none;padding:24px;border-radius:0 0 8px 8px;">
        <p>Hi ${toName}, a new leave request requires your approval.</p>
        ${buildTable(rows)}
        <p style="color:#666;font-size:13px;">Please action this request in Microsoft Teams.</p>
      </div>
    </div>`;

  try {
    await transporter.sendMail({
      from:    `"LeaveAgent" <${process.env.EMAIL_USER}>`,
      to:      toEmail,
      cc:      ccEmails.join(","),
      subject,
      html,
    });
    console.log(`[Email] Request sent -> To: ${toEmail} | CC: ${ccEmails.join(", ")}`);
  } catch (err) {
    console.warn(`[Email] Request notification failed:`, err);
  }
}

// ── Email 2: Decision (Approved / Rejected) ────────────────────────────────

export async function sendDecisionEmail(data: LeaveEmailData): Promise<void> {
  const hrEmail    = process.env.HR_EMAIL ?? "";
  const isApproved = data.status === "Approved";
  const statusText = isApproved ? "Approved" : "Rejected";
  const color      = isApproved ? "#107c10" : "#d13438";

  let ccEmails: string[];

  if (data.role === "employee") {
    // CC Manager + HR
    ccEmails = [data.managerEmail, hrEmail].filter(Boolean);
  } else {
    // Team Lead: CC HR only
    ccEmails = [hrEmail].filter(Boolean);
  }

  const subject = `Leave ${statusText}: ${data.employeeName} - ${getTypeLabel(data.requestType)} on ${data.displayDate}`;

  const rows: Array<[string, string]> = [
    ["Employee",    data.employeeName],
    ["Type",        getTypeLabel(data.requestType)],
    ["Date",        data.displayDate],
    ["Duration",    data.duration],
    ["Manager",     data.managerName],
    ...(data.role === "employee"
      ? [["Team Lead", data.teamleadName] as [string, string]]
      : []),
    ["Decision By", data.decidedBy ?? (data.role === "employee" ? data.teamleadName : data.managerName)],
    ["Status",      statusText],
  ];

  const html = `
    <div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto;">
      <div style="background:${color};padding:20px;border-radius:8px 8px 0 0;">
        <h2 style="color:white;margin:0;">Leave Request ${statusText}</h2>
      </div>
      <div style="border:1px solid #e0e0e0;border-top:none;padding:24px;border-radius:0 0 8px 8px;">
        <p>Hi ${data.employeeName}, your leave request has been
          <strong style="color:${color};">${statusText.toLowerCase()}</strong>
          by ${data.decidedBy}.
        </p>
        ${buildTable(rows)}
        <p style="color:${color};">
          ${isApproved
            ? "Your leave is confirmed. Please ensure handover before your leave date."
            : "Your request was not approved. Please speak with your approver for details."}
        </p>
        <p style="color:#666;font-size:13px;">This notification was sent by LeaveAgent.</p>
      </div>
    </div>`;

  try {
    await transporter.sendMail({
      from:    `"LeaveAgent" <${process.env.EMAIL_USER}>`,
      to:      data.employeeEmail,
      cc:      ccEmails.join(","),
      subject,
      html,
    });
    console.log(`[Email] Decision sent -> To: ${data.employeeEmail} | CC: ${ccEmails.join(", ")}`);
  } catch (err) {
    console.warn(`[Email] Decision email failed:`, err);
  }
}