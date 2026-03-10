import * as XLSX from "xlsx";
import * as fs from "fs";
import * as path from "path";

// ─────────────────────────────────────────────
// Types
// ─────────────────────────────────────────────

export interface Employee {
  name: string;
  email: string;
  manager: string;
  manager_email: string;
  teams_id?: string;
  manager_teams_id?: string;
}

export interface LeaveRecord {
  employee: string;
  email: string;
  type: string;
  date: string;
  end_date?: string;
  duration: string;
  status: "Pending" | "Approved" | "Rejected";
  approved_by?: string;
  requested_at: string;
  updated_at?: string;
}

export interface ConversationRef {
  userId: string;
  userName: string;
  conversationId: string;
  serviceUrl: string;
  tenantId?: string;
  botId: string;
}

// ─────────────────────────────────────────────
// In-memory conversation reference store
// ─────────────────────────────────────────────

const conversationRefs = new Map<string, ConversationRef>();

export function saveConversationRef(userId: string, ref: ConversationRef): void {
  conversationRefs.set(userId, ref);
  console.log(`[Excel] Saved conversation ref for ${ref.userName} (${userId})`);
}

export function getConversationRef(userId: string): ConversationRef | undefined {
  return conversationRefs.get(userId);
}

// ─────────────────────────────────────────────
// File paths (from env or defaults)
// ─────────────────────────────────────────────

const EMPLOYEES_PATH =
  process.env.EMPLOYEES_FILE_PATH ?? path.join(process.cwd(), "data", "Employees.xlsx");

const LEAVE_PATH =
  process.env.LEAVE_REQUESTS_FILE_PATH ??
  path.join(process.cwd(), "data", "LeaveRequests.xlsx");

const EMPLOYEE_HEADERS = [
  "name", "email", "manager", "manager_email", "teams_id", "manager_teams_id",
];

const LEAVE_HEADERS = [
  "employee", "email", "type", "date", "end_date",
  "duration", "status", "approved_by", "requested_at", "updated_at",
];

// ─────────────────────────────────────────────
// Helpers
// ─────────────────────────────────────────────

function ensureFile(filePath: string, headers: string[]): void {
  if (!fs.existsSync(filePath)) {
    const dir = path.dirname(filePath);
    if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet([headers]);
    XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
    XLSX.writeFile(wb, filePath);
    console.log(`[Excel] Created: ${filePath}`);
  }
}

function readSheet<T>(filePath: string, headers: string[]): T[] {
  ensureFile(filePath, headers);
  const wb = XLSX.readFile(filePath);
  const ws = wb.Sheets[wb.SheetNames[0]];
  return XLSX.utils.sheet_to_json<T>(ws, { defval: "" });
}

function writeSheet<T>(filePath: string, data: T[], headers: string[]): void {
  ensureFile(filePath, headers);
  const wb = XLSX.readFile(filePath);
  wb.Sheets[wb.SheetNames[0]] = XLSX.utils.json_to_sheet(data, { header: headers });
  XLSX.writeFile(wb, filePath);
}

// ─────────────────────────────────────────────
// Employees
// ─────────────────────────────────────────────

export function getAllEmployees(): Employee[] {
  return readSheet<Employee>(EMPLOYEES_PATH, EMPLOYEE_HEADERS);
}

export function findEmployee(nameOrEmail: string): Employee | undefined {
  const q = nameOrEmail.toLowerCase();
  return getAllEmployees().find(
    (e) => e.name?.toLowerCase() === q || e.email?.toLowerCase() === q
  );
}

// ─────────────────────────────────────────────
// Leave Requests
// ─────────────────────────────────────────────

export function getAllLeaveRequests(): LeaveRecord[] {
  return readSheet<LeaveRecord>(LEAVE_PATH, LEAVE_HEADERS);
}

export function addLeaveRequest(record: LeaveRecord): void {
  const all = getAllLeaveRequests();
  all.push(record);
  writeSheet(LEAVE_PATH, all, LEAVE_HEADERS);
  console.log(`[Excel] ✅ Added: ${record.employee} - ${record.type} on ${record.date}`);
}

export function updateLeaveStatus(
  employee: string,
  date: string,
  status: "Approved" | "Rejected",
  approvedBy: string
): boolean {
  const all = getAllLeaveRequests();
  const idx = all.findIndex(
    (r) =>
      r.employee?.toLowerCase() === employee.toLowerCase() &&
      r.date === date &&
      r.status === "Pending"
  );
  if (idx === -1) return false;

  all[idx].status = status;
  all[idx].approved_by = approvedBy;
  all[idx].updated_at = new Date().toISOString();
  writeSheet(LEAVE_PATH, all, LEAVE_HEADERS);
  console.log(`[Excel] ✅ Updated ${employee} → ${status}`);
  return true;
}

export function isDuplicateRequest(employee: string, date: string): boolean {
  return getAllLeaveRequests().some(
    (r) =>
      r.employee?.toLowerCase() === employee.toLowerCase() &&
      r.date === date &&
      r.status !== "Rejected"
  );
}

export function getTodaysAbsences(): LeaveRecord[] {
  const today = new Date().toISOString().split("T")[0];
  return getAllLeaveRequests().filter((r) => r.date === today && r.status === "Approved");
}