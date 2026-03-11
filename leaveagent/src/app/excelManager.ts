/**
 * excelManager.ts
 * Handles all Excel read/write operations and in-memory conversation refs.
 */

import * as XLSX from "xlsx";
import * as path from "path";
import * as fs   from "fs";

const DATA_DIR      = path.join(process.cwd(), "data");
const EMPLOYEES_PATH = path.join(DATA_DIR, "Employees.xlsx");
const LEAVE_PATH     = path.join(DATA_DIR, "LeaveRequests.xlsx");

// ── Types ──────────────────────────────────────────────────────────────────

export interface Employee {
  name:              string;
  email:             string;
  manager:           string;
  manager_email:     string;
  manager_teams_id:  string;
  teamlead:          string;       // ADDED: per-employee team lead name
  teamlead_email:    string;       // ADDED: per-employee team lead email
  teams_id:          string;
}

export interface LeaveRecord {
  employee:     string;
  email:        string;
  type:         string;
  date:         string;
  end_date?:    string;
  duration:     string;
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

// ── In-memory conversation refs ────────────────────────────────────────────

const conversationRefs = new Map<string, ConversationRef>();

export function saveConversationRef(userId: string, ref: ConversationRef): void {
  conversationRefs.set(userId, ref);
  console.log(`[Excel] Saved conversation ref for ${ref.userName} (${userId})`);
}

export function getConversationRef(userId: string): ConversationRef | undefined {
  return conversationRefs.get(userId);
}

// ── Employees ──────────────────────────────────────────────────────────────

export function getAllEmployees(): Employee[] {
  try {
    const wb   = XLSX.readFile(EMPLOYEES_PATH);
    const ws   = wb.Sheets[wb.SheetNames[0]];
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
        header: ["employee","email","type","date","end_date","duration","status","approved_by","requested_at","updated_at"],
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
    header: ["employee","email","type","date","end_date","duration","status","approved_by","requested_at","updated_at"],
  });
  XLSX.writeFile(wb, LEAVE_PATH);
  console.log(`[Excel] Added: ${record.employee} - ${record.type} on ${record.date}`);
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
    header: ["employee","email","type","date","end_date","duration","status","approved_by","requested_at","updated_at"],
  });
  XLSX.writeFile(wb, LEAVE_PATH);
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