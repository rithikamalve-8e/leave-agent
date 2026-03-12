import * as XLSX from "xlsx";
import * as path from "path";
import * as fs   from "fs";

const DATA_DIR       = path.join(process.cwd(), "data");
const EMPLOYEES_PATH = path.join(DATA_DIR, "Employees.xlsx");
const LEAVE_PATH     = path.join(DATA_DIR, "LeaveRequests.xlsx");

// ── Types ──────────────────────────────────────────────────────────────────

export interface Employee {
  name:               string;
  email:              string;
  role:               "employee" | "teamlead";
  manager:            string;
  manager_email:      string;
  manager_teams_id:   string;
  teamlead:           string;
  teamlead_email:     string;
  teamlead_teams_id:  string;
  teams_id:           string;
  leave_balance:      number; // annual leave days remaining (from Employees.xlsx)
}

export interface LeaveRecord {
  employee:     string;
  email:        string;
  type:         string;
  date:         string;
  end_date?:    string;
  duration:     string;
  days_count:   number;  // ADDED: number of working days in the request
  reason?:      string;  // ADDED: optional reason
  status:       string;
  approved_by?: string;
  requested_at: string;
  updated_at?:  string;
}

export interface ConversationRef {
  userId:         string;
  userName:       string;
  conversationId: string;
  serviceUrl:     string;
  tenantId?:      string;
  botId:          string;
}

// ── Leave balance check result ─────────────────────────────────────────────

export interface LeaveBalanceResult {
  requested:  number;  // days requested
  balance:    number;  // days available
  granted:    number;  // days that will be approved
  lop:        number;  // days that will be Loss of Pay
  hasLop:     boolean; // true if any days are LOP
}

// ── In-memory conversation refs ────────────────────────────────────────────

const conversationRefs = new Map<string, ConversationRef>();

export function saveConversationRef(userId: string, ref: ConversationRef): void {
  conversationRefs.set(userId, ref);
  console.log(`[Excel] Saved conversation ref for ${ref.userName} (${userId})`);
}

export function getConversationRef(userId: string): ConversationRef | undefined {
  return conversationRefs.get(userId);
}

// ── Working days counter ───────────────────────────────────────────────────
// Counts Mon–Fri days between start and end date (inclusive)

export function countWorkingDays(startDate: string, endDate?: string): number {
  if (!startDate) return 0;
  const start = new Date(startDate + "T00:00:00");
  const end   = endDate ? new Date(endDate + "T00:00:00") : new Date(startDate + "T00:00:00");

  let count = 0;
  const current = new Date(start);
  while (current <= end) {
    const day = current.getDay();
    if (day !== 0 && day !== 6) count++; // skip Saturday (6) and Sunday (0)
    current.setDate(current.getDate() + 1);
  }
  return count;
}

// ── Leave balance check ────────────────────────────────────────────────────
// Only applies to LEAVE type — WFH and SICK don't consume leave balance

export function checkLeaveBalance(
  employee: Employee,
  daysRequested: number,
  leaveType: string
): LeaveBalanceResult {
  // WFH and SICK don't consume leave balance
  if (leaveType !== "LEAVE") {
    return { requested: daysRequested, balance: 999, granted: daysRequested, lop: 0, hasLop: false };
  }

  const balance = Number(employee.leave_balance ?? 0);
  const granted = Math.min(daysRequested, balance);
  const lop     = Math.max(0, daysRequested - balance);

  return {
    requested: daysRequested,
    balance,
    granted,
    lop,
    hasLop: lop > 0,
  };
}

// ── Deduct leave balance after approval ───────────────────────────────────

export function deductLeaveBalance(employeeName: string, days: number): void {
  try {
    const wb   = XLSX.readFile(EMPLOYEES_PATH);
    const ws   = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json<Employee>(ws);

    const idx = rows.findIndex((e) => e.name?.toLowerCase() === employeeName.toLowerCase());
    if (idx === -1) return;

    rows[idx].leave_balance = Math.max(0, Number(rows[idx].leave_balance ?? 0) - days);

    wb.Sheets[wb.SheetNames[0]] = XLSX.utils.json_to_sheet(rows, {
      header: ["name","email","role","manager","manager_email","manager_teams_id",
               "teamlead","teamlead_email","teamlead_teams_id","teams_id","leave_balance"],
    });
    XLSX.writeFile(wb, EMPLOYEES_PATH);
    console.log(`[Excel] Deducted ${days} leave days from ${employeeName}. Remaining: ${rows[idx].leave_balance}`);
  } catch (err) {
    console.warn(`[Excel] Failed to deduct leave balance:`, err);
  }
}

// ── Employees ──────────────────────────────────────────────────────────────

export function getAllEmployees(): Employee[] {
  try {
    const wb = XLSX.readFile(EMPLOYEES_PATH);
    const ws = wb.Sheets[wb.SheetNames[0]];
    return XLSX.utils.sheet_to_json<Employee>(ws);
  } catch {
    return [];
  }
}

export function findEmployee(nameOrEmail: string): Employee | undefined {
  const all = getAllEmployees();
  const key = nameOrEmail.toLowerCase();
  return all.find(
    (e) =>
      e.name?.toLowerCase()  === key ||
      e.email?.toLowerCase() === key
  );
}

// ── Leave Requests ─────────────────────────────────────────────────────────

function ensureLeaveFile(): void {
  if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });
  if (!fs.existsSync(LEAVE_PATH)) {
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(
      wb,
      XLSX.utils.json_to_sheet([], {
        header: ["employee","email","type","date","end_date","duration","days_count","reason","status","approved_by","requested_at","updated_at"],
      }),
      "LeaveRequests"
    );
    XLSX.writeFile(wb, LEAVE_PATH);
    console.log(`[Excel] Created: ${LEAVE_PATH}`);
  }
}

export function getAllLeaveRequests(): LeaveRecord[] {
  ensureLeaveFile();
  const wb = XLSX.readFile(LEAVE_PATH);
  const ws = wb.Sheets[wb.SheetNames[0]];
  return XLSX.utils.sheet_to_json<LeaveRecord>(ws);
}

export function addLeaveRequest(record: LeaveRecord): void {
  ensureLeaveFile();
  const wb      = XLSX.readFile(LEAVE_PATH);
  const ws      = wb.Sheets[wb.SheetNames[0]];
  const records = XLSX.utils.sheet_to_json<LeaveRecord>(ws);
  records.push(record);
  wb.Sheets[wb.SheetNames[0]] = XLSX.utils.json_to_sheet(records, {
    header: ["employee","email","type","date","end_date","duration","days_count","reason","status","approved_by","requested_at","updated_at"],
  });
  XLSX.writeFile(wb, LEAVE_PATH);
  console.log(`[Excel] Added: ${record.employee} - ${record.type} on ${record.date} (${record.days_count} days)`);
}

export function updateLeaveStatus(
  employee: string,
  date: string,
  status: string,
  approvedBy: string
): boolean {
  ensureLeaveFile();
  const wb      = XLSX.readFile(LEAVE_PATH);
  const ws      = wb.Sheets[wb.SheetNames[0]];
  const records = XLSX.utils.sheet_to_json<LeaveRecord>(ws);

  const idx = records.findIndex(
    (r) =>
      r.employee?.toLowerCase() === employee.toLowerCase() &&
      r.date === date &&
      r.status === "Pending"
  );

  if (idx === -1) return false;

  records[idx].status      = status;
  records[idx].approved_by = approvedBy;
  records[idx].updated_at  = new Date().toISOString();

  wb.Sheets[wb.SheetNames[0]] = XLSX.utils.json_to_sheet(records, {
    header: ["employee","email","type","date","end_date","duration","days_count","reason","status","approved_by","requested_at","updated_at"],
  });
  XLSX.writeFile(wb, LEAVE_PATH);

  // Deduct leave balance on approval for LEAVE type
  if (status === "Approved" && records[idx].type === "LEAVE") {
    deductLeaveBalance(employee, records[idx].days_count ?? 1);
  }

  return true;
}

export function isDuplicateRequest(employee: string, date: string): boolean {
  const records = getAllLeaveRequests();
  return records.some(
    (r) =>
      r.employee?.toLowerCase() === employee.toLowerCase() &&
      r.date === date &&
      r.status === "Pending"
  );
}

export function getTodaysAbsences(): LeaveRecord[] {
  const today   = new Date().toISOString().split("T")[0];
  const records = getAllLeaveRequests();
  return records.filter((r) => r.date === today && r.status === "Approved");
}