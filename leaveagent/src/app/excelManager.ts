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
  leave_balance:      number;
}
 
export interface LeaveRecord {
  employee:     string;
  email:        string;
  type:         string;
  date:         string;
  end_date?:    string;
  duration:     string;
  days_count:   number;
  reason?:      string;
  status:       string;
  approved_by?: string;
  requested_at: string;
  updated_at?:  string;
}
 
export interface MonthlySummaryRecord {
  Month:      string;  // MM/YYYY
  Emp:        string;
  Leaves:     number;
  WFH:        number;
  unapproved: number;
  approved:   number;
}
 
export interface ConversationRef {
  userId:         string;
  userName:       string;
  conversationId: string;
  serviceUrl:     string;
  tenantId?:      string;
  botId:          string;
}
 
export interface LeaveBalanceResult {
  requested:  number;
  balance:    number;
  granted:    number;
  lop:        number;
  hasLop:     boolean;
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
 
export function countWorkingDays(startDate: string, endDate?: string): number {
  if (!startDate) return 0;
  const start = new Date(startDate + "T00:00:00");
  const end   = endDate ? new Date(endDate + "T00:00:00") : new Date(startDate + "T00:00:00");
 
  let count = 0;
  const current = new Date(start);
  while (current <= end) {
    const day = current.getDay();
    if (day !== 0 && day !== 6) count++;
    current.setDate(current.getDate() + 1);
  }
  return count;
}
 
// ── Leave balance check ────────────────────────────────────────────────────
 
export function checkLeaveBalance(
  employee: Employee,
  daysRequested: number,
  leaveType: string
): LeaveBalanceResult {
  const BALANCE_CONSUMING_TYPES = ["LEAVE", "SICK", "MATERNITY", "PATERNITY", "ADOPTION", "MARRIAGE"];
  if (!BALANCE_CONSUMING_TYPES.includes(leaveType?.toUpperCase())) {
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
 
// ── Leave file setup ───────────────────────────────────────────────────────
 
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
    XLSX.utils.book_append_sheet(
      wb,
      XLSX.utils.json_to_sheet([], {
        header: ["Month","Emp","Leaves","WFH","unapproved","approved"],
      }),
      "MonthlySummary"
    );
    XLSX.writeFile(wb, LEAVE_PATH);
    console.log(`[Excel] Created: ${LEAVE_PATH}`);
  } else {
    // File exists — ensure MonthlySummary sheet is present
    const wb = XLSX.readFile(LEAVE_PATH);
    if (!wb.SheetNames.includes("MonthlySummary")) {
      XLSX.utils.book_append_sheet(
        wb,
        XLSX.utils.json_to_sheet([], {
          header: ["Month","Emp","Leaves","WFH","unapproved","approved"],
        }),
        "MonthlySummary"
      );
      XLSX.writeFile(wb, LEAVE_PATH);
      console.log(`[Excel] Added MonthlySummary sheet to existing file`);
    }
  }
}
 
// ── Monthly Summary helpers ────────────────────────────────────────────────
 
function getMonthlySummaryRecords(wb: XLSX.WorkBook): MonthlySummaryRecord[] {
  const ws = wb.Sheets["MonthlySummary"];
  if (!ws) return [];
  return XLSX.utils.sheet_to_json<MonthlySummaryRecord>(ws);
}
 
function writeMonthlySummaryRecords(wb: XLSX.WorkBook, records: MonthlySummaryRecord[]): void {
  wb.Sheets["MonthlySummary"] = XLSX.utils.json_to_sheet(records, {
    header: ["Month","Emp","Leaves","WFH","unapproved","approved"],
  });
}
 
function updateMonthlySummary(
  employeeName: string,
  leaveType:    string,
  date:         string,
  daysCount:    number,
  status:       "Approved" | "Rejected"
): void {
  try {
    ensureLeaveFile();
    const wb      = XLSX.readFile(LEAVE_PATH);
    const records = getMonthlySummaryRecords(wb);
 
    const d     = new Date(date + "T00:00:00");
    const month = `${String(d.getMonth() + 1).padStart(2, "0")}/${d.getFullYear()}`;
 
    // Count current pending days for this employee this month
    const allLeave    = getAllLeaveRequests();
    const pendingDays = allLeave
      .filter((r) => {
        const rd = new Date(r.date + "T00:00:00");
        const rm = `${String(rd.getMonth() + 1).padStart(2, "0")}/${rd.getFullYear()}`;
        return (
          r.employee?.toLowerCase() === employeeName.toLowerCase() &&
          rm === month &&
          r.status === "Pending"
        );
      })
      .reduce((sum, r) => sum + (Number(r.days_count) || 1), 0);
 
    const idx = records.findIndex(
      (r) =>
        r.Month === month &&
        r.Emp?.toLowerCase() === employeeName.toLowerCase()
    );
 
    if (idx === -1) {
      const newRow: MonthlySummaryRecord = {
        Month:      month,
        Emp:        employeeName,
        Leaves:     0,
        WFH:        0,
        unapproved: pendingDays,
        approved:   0,
      };
 
      if (status === "Approved") {
        if (leaveType === "WFH") {
          newRow.WFH    = daysCount;
        } else {
          newRow.Leaves = daysCount;
        }
        newRow.approved = daysCount;
      }
 
      records.push(newRow);
    } else {
      // Update unapproved = current pending days (live count)
      records[idx].unapproved = pendingDays;
 
      if (status === "Approved") {
        if (leaveType === "WFH") {
          records[idx].WFH    = (Number(records[idx].WFH)    || 0) + daysCount;
        } else {
          records[idx].Leaves = (Number(records[idx].Leaves) || 0) + daysCount;
        }
        records[idx].approved = (Number(records[idx].approved) || 0) + daysCount;
      }
    }
 
    writeMonthlySummaryRecords(wb, records);
    XLSX.writeFile(wb, LEAVE_PATH);
    console.log(`[Excel] MonthlySummary updated for ${employeeName} - ${month} | pending: ${pendingDays}`);
  } catch (err) {
    console.warn(`[Excel] Failed to update MonthlySummary:`, err);
  }
}
 
// ── Leave Requests ─────────────────────────────────────────────────────────
 
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
 
  // Deduct leave balance on approval for non-WFH types
  if (status === "Approved" && records[idx].type !== "WFH") {
    deductLeaveBalance(employee, records[idx].days_count ?? 1);
  }
 
  // Update monthly summary on approval or rejection
  if (status === "Approved" || status === "Rejected") {
    updateMonthlySummary(
      employee,
      records[idx].type,
      records[idx].date,
      records[idx].days_count ?? 1,
      status as "Approved" | "Rejected"
    );
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