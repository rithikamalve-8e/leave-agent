import * as XLSX from "xlsx";
import * as path from "path";
import * as fs   from "fs";

const DATA_DIR       = path.join(process.cwd(), "data");
const EMPLOYEES_PATH = path.join(DATA_DIR, "Employees.xlsx");
const LEAVE_PATH     = path.join(DATA_DIR, "LeaveRequests.xlsx");

const EMPLOYEE_HEADERS = [
  "name","email","role","manager","manager_email","manager_teams_id",
  "teamlead","teamlead_email","teamlead_teams_id","teams_id",
  "leave_balance","carry_forward","year_entitlement_start"
];

const LEAVE_HEADERS = [
  "employee","email","type","date","end_date","duration",
  "days_count","lop_days","reason","status","approved_by","requested_at","updated_at"
];

const SUMMARY_HEADERS = [
  "month","employee","opening","available","leaves","wfh","lop","closing","pending"
];

// ── Types ──────────────────────────────────────────────────────────────────

export interface Employee {
  name:                    string;
  email:                   string;
  role:                    "employee" | "teamlead";
  manager:                 string;
  manager_email:           string;
  manager_teams_id:        string;
  teamlead:                string;
  teamlead_email:          string;
  teamlead_teams_id:       string;
  teams_id:                string;
  leave_balance:           number;
  carry_forward:           number;  // days from last year (max 6, constant all year)
  year_entitlement_start:  number;  // YYYY — prevents double Jan 1st credit
}

export interface LeaveRecord {
  employee:      string;
  email:         string;
  type:          string;
  date:          string;
  end_date?:     string;
  duration:      string;
  days_count:    number;
  lop_days:      number;   // locked at submission time
  reason?:       string;
  status:        string;
  approved_by?:  string;
  requested_at:  string;
  updated_at?:   string;
}

export interface MonthlySummaryRecord {
  month:     string;  // MM/YYYY
  employee:  string;
  opening:   number;
  available: number;
  leaves:    number;
  wfh:       number;
  lop:       number;
  closing:   number;
  pending:   number;
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
  splits?:    MonthlyLopSplit[];
}

export interface MonthlyLopSplit {
  month:    string;  // MM/YYYY
  days:     number;
  balance:  number;
  lop:      number;
  granted:  number;
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
  const start   = new Date(startDate + "T00:00:00");
  const end     = endDate ? new Date(endDate + "T00:00:00") : new Date(startDate + "T00:00:00");
  let count     = 0;
  const current = new Date(start);
  while (current <= end) {
    const day = current.getDay();
    if (day !== 0 && day !== 6) count++;
    current.setDate(current.getDate() + 1);
  }
  return count;
}

// ── Core balance formula ───────────────────────────────────────────────────
// balance = carry_forward + (currentMonth × 1.5) - totalUsedThisYear + totalLOPThisYear

export function getLeaveBalance(
  employeeName: string,
  asOfDate?:    Date       // defaults to today
): number {
  const now          = asOfDate ?? new Date();
  const currentMonth = now.getMonth() + 1;  // 1-12
  const currentYear  = now.getFullYear();

  const emp = findEmployee(employeeName);
  if (!emp) return 0;

  const carry = Number(emp.carry_forward ?? 0);

  // Get all approved non-WFH leaves for this employee this year
  const allRequests = getAllLeaveRequests();
  const approvedThisYear = allRequests.filter((r) => {
    const d = new Date(r.date + "T00:00:00");
    return (
      r.employee?.toLowerCase() === employeeName.toLowerCase() &&
      r.status === "Approved" &&
      r.type   !== "WFH" &&
      d.getFullYear() === currentYear
    );
  });

  const totalUsed = approvedThisYear.reduce((sum, r) => sum + (Number(r.days_count) || 0), 0);
  const totalLOP  = approvedThisYear.reduce((sum, r) => sum + (Number(r.lop_days)   || 0), 0);

  const yearEntitlement = currentMonth * 1.5;
  const balance         = carry + yearEntitlement - totalUsed + totalLOP;

  return Math.max(0, balance);
}

// ── Leave balance check ────────────────────────────────────────────────────

export function checkLeaveBalance(
  employee:      Employee,
  daysRequested: number,
  leaveType:     string,
  startDate?:    string,
  endDate?:      string
): LeaveBalanceResult {
  const BALANCE_CONSUMING = [
    "LEAVE","SICK","MATERNITY","PATERNITY","ADOPTION","MARRIAGE"
  ];

  // WFH is free — no balance consumed
  if (!BALANCE_CONSUMING.includes(leaveType?.toUpperCase())) {
    return {
      requested: daysRequested,
      balance:   999,
      granted:   daysRequested,
      lop:       0,
      hasLop:    false
    };
  }

  // Cross-month check
  if (startDate && endDate && startDate !== endDate) {
    const startMonth = startDate.substring(0, 7); // "2026-02"
    const endMonth   = endDate.substring(0, 7);   // "2026-03"

    if (startMonth !== endMonth) {
      const result = calculateCrossMonthLop(employee, startDate, endDate, leaveType);
      return {
        requested: result.totalDays,
        balance:   getLeaveBalance(employee.name),
        granted:   result.totalGranted,
        lop:       result.totalLop,
        hasLop:    result.totalLop > 0,
        splits:    result.splits
      };
    }
  }

  // Single month
  const balance = getLeaveBalance(employee.name);
  const granted = Math.min(daysRequested, balance);
  const lop     = Math.max(0, daysRequested - balance);

  return { requested: daysRequested, balance, granted, lop, hasLop: lop > 0 };
}

// ── Cross-month LOP split ──────────────────────────────────────────────────

export function calculateCrossMonthLop(
  employee:  Employee,
  startDate: string,
  endDate:   string,
  leaveType: string
): {
  splits:       MonthlyLopSplit[];
  totalDays:    number;
  totalLop:     number;
  totalGranted: number;
} {
  const BALANCE_CONSUMING = [
    "LEAVE","SICK","MATERNITY","PATERNITY","ADOPTION","MARRIAGE"
  ];
  if (!BALANCE_CONSUMING.includes(leaveType?.toUpperCase())) {
    const days = countWorkingDays(startDate, endDate);
    return { splits: [], totalDays: days, totalLop: 0, totalGranted: days };
  }

  // Group working days by month
  const start       = new Date(startDate + "T00:00:00");
  const end         = new Date(endDate   + "T00:00:00");
  const current     = new Date(start);
  const monthDays:  Record<string, number> = {};
  const monthOrder: string[] = [];

  while (current <= end) {
    const day = current.getDay();
    if (day !== 0 && day !== 6) {
      const key = `${String(current.getMonth() + 1).padStart(2, "0")}/${current.getFullYear()}`;
      if (!monthDays[key]) {
        monthDays[key] = 0;
        monthOrder.push(key);
      }
      monthDays[key]++;
    }
    current.setDate(current.getDate() + 1);
  }

  // Request submission month
  const requestDate     = new Date();
  const requestMonthNum = requestDate.getMonth() + 1;
  const requestYear     = requestDate.getFullYear();

  // Current balance at time of request
  let runningBalance = getLeaveBalance(employee.name);
  const splits: MonthlyLopSplit[] = [];

  for (const monthKey of monthOrder) {
    const [mm, yyyy] = monthKey.split("/").map(Number);
    const days       = monthDays[monthKey];

    // Future month → no balance available yet (accrual hasn't happened)
    const isFutureMonth =
      yyyy > requestYear ||
      (yyyy === requestYear && mm > requestMonthNum);

    const balanceThisMonth = isFutureMonth ? 0 : runningBalance;
    const granted          = Math.min(days, balanceThisMonth);
    const lop              = Math.max(0, days - balanceThisMonth);

    splits.push({ month: monthKey, days, balance: balanceThisMonth, lop, granted });

    // Carry remaining to next month
    runningBalance = Math.max(0, balanceThisMonth - days);
  }

  const totalDays    = splits.reduce((s, r) => s + r.days,    0);
  const totalLop     = splits.reduce((s, r) => s + r.lop,     0);
  const totalGranted = splits.reduce((s, r) => s + r.granted, 0);

  return { splits, totalDays, totalLop, totalGranted };
}

// ── Save employees ─────────────────────────────────────────────────────────

function saveEmployees(rows: Employee[]): void {
  const wb = XLSX.readFile(EMPLOYEES_PATH);
  wb.Sheets[wb.SheetNames[0]] = XLSX.utils.json_to_sheet(rows, { header: EMPLOYEE_HEADERS });
  XLSX.writeFile(wb, EMPLOYEES_PATH);
}

// ── Deduct leave balance ───────────────────────────────────────────────────
// NOTE: leave_balance column in Employees.xlsx is now only used
// for HR manual adjustments and initial setup.
// Actual computed balance comes from getLeaveBalance() formula.
// deductLeaveBalance is kept for backward compatibility during
// Excel phase — will be removed when migrating to Postgres.

export function deductLeaveBalance(employeeName: string, days: number): void {
  try {
    const wb   = XLSX.readFile(EMPLOYEES_PATH);
    const ws   = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json<Employee>(ws);
    const idx  = rows.findIndex((e) => e.name?.toLowerCase() === employeeName.toLowerCase());
    if (idx === -1) return;
    rows[idx].leave_balance = Math.max(0, Number(rows[idx].leave_balance ?? 0) - days);
    saveEmployees(rows);
    console.log(`[Excel] Deducted ${days} days from ${employeeName}. Stored balance: ${rows[idx].leave_balance}`);
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

// ── Jan 1st carry forward job ──────────────────────────────────────────────

export function runYearStartAccrual(): void {
  try {
    const now         = new Date();
    const currentYear = now.getFullYear();

    const wb   = XLSX.readFile(EMPLOYEES_PATH);
    const ws   = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json<Employee>(ws);

    let changed = false;

    for (let i = 0; i < rows.length; i++) {
      const emp = rows[i];

      // Already processed this year → skip
      if (Number(emp.year_entitlement_start) === currentYear) {
        console.log(`[Accrual] ${emp.name} already processed for ${currentYear}`);
        continue;
      }

      // carry_forward = MIN(current balance, 6)
      // current balance at Dec 31 = whatever is left
      const prevBalance   = Number(emp.leave_balance ?? 0);
      const carryForward  = Math.min(prevBalance, 6);

      rows[i].carry_forward          = carryForward;
      rows[i].year_entitlement_start = currentYear;
      changed = true;

      console.log(
        `[Accrual] ${emp.name}: carry_forward=${carryForward} for ${currentYear}` +
        ` (prev balance was ${prevBalance})`
      );
    }

    if (changed) {
      saveEmployees(rows);
      console.log(`[Accrual] Year start accrual saved`);
    } else {
      console.log(`[Accrual] Year start already processed`);
    }
  } catch (err) {
    console.warn(`[Accrual] Year start accrual failed:`, err);
  }
}

// ── Leave file setup ───────────────────────────────────────────────────────

function ensureLeaveFile(): void {
  if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });
  if (!fs.existsSync(LEAVE_PATH)) {
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(
      wb,
      XLSX.utils.json_to_sheet([], { header: LEAVE_HEADERS }),
      "LeaveRequests"
    );
    XLSX.utils.book_append_sheet(
      wb,
      XLSX.utils.json_to_sheet([], { header: SUMMARY_HEADERS }),
      "MonthlySummary"
    );
    XLSX.writeFile(wb, LEAVE_PATH);
    console.log(`[Excel] Created: ${LEAVE_PATH}`);
  } else {
    const wb = XLSX.readFile(LEAVE_PATH);
    if (!wb.SheetNames.includes("MonthlySummary")) {
      XLSX.utils.book_append_sheet(
        wb,
        XLSX.utils.json_to_sheet([], { header: SUMMARY_HEADERS }),
        "MonthlySummary"
      );
      XLSX.writeFile(wb, LEAVE_PATH);
      console.log(`[Excel] Added MonthlySummary sheet`);
    }
  }
}

// ── Monthly Summary helpers ────────────────────────────────────────────────

function getMonthlySummaryRecords(wb: XLSX.WorkBook): MonthlySummaryRecord[] {
  const ws = wb.Sheets["MonthlySummary"];
  if (!ws) return [];
  return XLSX.utils.sheet_to_json<MonthlySummaryRecord>(ws);
}

function writeMonthlySummaryRecords(
  wb:      XLSX.WorkBook,
  records: MonthlySummaryRecord[]
): void {
  wb.Sheets["MonthlySummary"] = XLSX.utils.json_to_sheet(records, {
    header: SUMMARY_HEADERS
  });
}

// ── Build monthly summary for ONE employee ONE month ──────────────────────
// Called by scheduler at end of month for every employee

export function buildEmployeeMonthlySummary(
  employeeName: string,
  monthNum:     number,   // 1-12
  year:         number
): MonthlySummaryRecord {
  const monthKey = `${String(monthNum).padStart(2, "0")}/${year}`;

  const emp   = findEmployee(employeeName);
  const carry = Number(emp?.carry_forward ?? 0);

  // All approved non-WFH requests for this employee this year
  const allRequests    = getAllLeaveRequests();
  const approvedNonWFH = allRequests.filter((r) => {
    const d = new Date(r.date + "T00:00:00");
    return (
      r.employee?.toLowerCase() === employeeName.toLowerCase() &&
      r.status === "Approved" &&
      r.type   !== "WFH" &&
      d.getFullYear() === year
    );
  });

  // Split into: before this month vs this month
  const beforeMonth = approvedNonWFH.filter((r) => {
    const d = new Date(r.date + "T00:00:00");
    return d.getMonth() + 1 < monthNum;
  });

  const thisMonth = approvedNonWFH.filter((r) => {
    const d = new Date(r.date + "T00:00:00");
    return d.getMonth() + 1 === monthNum;
  });

  // Totals before this month
  const totalUsedBeforeM = beforeMonth.reduce((s, r) => s + (Number(r.days_count) || 0), 0);
  const totalLOPBeforeM  = beforeMonth.reduce((s, r) => s + (Number(r.lop_days)   || 0), 0);

  // This month totals
  const leavesThisMonth  = thisMonth.reduce((s, r) => s + (Number(r.days_count) || 0), 0);
  const lopThisMonth     = thisMonth.reduce((s, r) => s + (Number(r.lop_days)   || 0), 0);

  // WFH this month
  const wfhThisMonth = allRequests
    .filter((r) => {
      const d = new Date(r.date + "T00:00:00");
      return (
        r.employee?.toLowerCase() === employeeName.toLowerCase() &&
        r.status === "Approved" &&
        r.type   === "WFH" &&
        d.getMonth() + 1 === monthNum &&
        d.getFullYear()   === year
      );
    })
    .reduce((s, r) => s + (Number(r.days_count) || 0), 0);

  // Pending snapshot at end of month
  const pendingThisMonth = allRequests
    .filter((r) => {
      const d = new Date(r.date + "T00:00:00");
      return (
        r.employee?.toLowerCase() === employeeName.toLowerCase() &&
        r.status === "Pending" &&
        d.getMonth() + 1 === monthNum &&
        d.getFullYear()   === year
      );
    })
    .reduce((s, r) => s + (Number(r.days_count) || 0), 0);

  // Core formula
  // opening   = carry + ((M-1) × 1.5) - totalUsed(Jan→M-1) + totalLOP(Jan→M-1)
  // available = opening + 1.5
  // closing   = available - leaves + lop
  const opening   = carry + ((monthNum - 1) * 1.5) - totalUsedBeforeM + totalLOPBeforeM;
  const available = opening + 1.5;
  const closing   = available - leavesThisMonth + lopThisMonth;

  return {
    month:     monthKey,
    employee:  employeeName,
    opening:   Math.max(0, opening),
    available: Math.max(0, available),
    leaves:    leavesThisMonth,
    wfh:       wfhThisMonth,
    lop:       lopThisMonth,
    closing:   Math.max(0, closing),
    pending:   pendingThisMonth,
  };
}

// ── Build and save monthly summary for ALL employees for ONE month ─────────
// Called by scheduler at end of month

export function buildAndSaveMonthlySummary(
  monthNum: number,  // 1-12
  year:     number
): void {
  try {
    ensureLeaveFile();
    const employees = getAllEmployees();
    const wb        = XLSX.readFile(LEAVE_PATH);
    const records   = getMonthlySummaryRecords(wb);

    const monthKey = `${String(monthNum).padStart(2, "0")}/${year}`;

    for (const emp of employees) {
      // Calculate this employee's summary for this month
      const summary = buildEmployeeMonthlySummary(emp.name, monthNum, year);

      // Check if row already exists (re-run safety)
      const idx = records.findIndex(
        (r) =>
          r.month    === monthKey &&
          r.employee?.toLowerCase() === emp.name.toLowerCase()
      );

      if (idx === -1) {
        // Insert new row
        records.push(summary);
        console.log(`[Summary] Inserted: ${emp.name} ${monthKey}`);
      } else {
        // Overwrite existing row (re-run safe)
        records[idx] = summary;
        console.log(`[Summary] Updated: ${emp.name} ${monthKey}`);
      }
    }

    // Sort by month then employee for clean Excel view
    records.sort((a, b) => {
      if (a.month === b.month) return a.employee.localeCompare(b.employee);
      const [am, ay] = a.month.split("/").map(Number);
      const [bm, by] = b.month.split("/").map(Number);
      return ay !== by ? ay - by : am - bm;
    });

    writeMonthlySummaryRecords(wb, records);
    XLSX.writeFile(wb, LEAVE_PATH);
    console.log(`[Summary] Monthly summary built for ${monthKey} — ${employees.length} employees`);
  } catch (err) {
    console.warn(`[Summary] Failed to build monthly summary:`, err);
  }
}

// ── Get monthly summary for HR report ─────────────────────────────────────

export function getMonthlySummaryForMonth(monthKey: string): MonthlySummaryRecord[] {
  ensureLeaveFile();
  const wb      = XLSX.readFile(LEAVE_PATH);
  const records = getMonthlySummaryRecords(wb);
  return records.filter((r) => r.month === monthKey);
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
    header: LEAVE_HEADERS
  });
  XLSX.writeFile(wb, LEAVE_PATH);
  console.log(`[Excel] Added: ${record.employee} - ${record.type} on ${record.date} (${record.days_count} days, lop: ${record.lop_days})`);
}

export function updateLeaveStatus(
  employee:   string,
  date:       string,
  status:     string,
  approvedBy: string
): boolean {
  ensureLeaveFile();
  const wb      = XLSX.readFile(LEAVE_PATH);
  const ws      = wb.Sheets[wb.SheetNames[0]];
  const records = XLSX.utils.sheet_to_json<LeaveRecord>(ws);

  const idx = records.findIndex(
    (r) =>
      r.employee?.toLowerCase() === employee.toLowerCase() &&
      r.date   === date &&
      r.status === "Pending"
  );

  if (idx === -1) return false;

  records[idx].status      = status;
  records[idx].approved_by = approvedBy;
  records[idx].updated_at  = new Date().toISOString();

  wb.Sheets[wb.SheetNames[0]] = XLSX.utils.json_to_sheet(records, {
    header: LEAVE_HEADERS
  });
  XLSX.writeFile(wb, LEAVE_PATH);

  // Deduct from stored leave_balance (used for HR manual tracking)
  // getLeaveBalance() formula is the source of truth for computed balance
  if (status === "Approved" && records[idx].type !== "WFH") {
    deductLeaveBalance(employee, records[idx].days_count ?? 1);
  }

  // Monthly summary is NO LONGER updated here
  // It is built by the end-of-month scheduler (buildAndSaveMonthlySummary)
  console.log(`[Excel] Status updated: ${employee} ${date} → ${status} by ${approvedBy}`);

  return true;
}

export function isDuplicateRequest(employee: string, date: string): boolean {
  const records = getAllLeaveRequests();
  return records.some(
    (r) =>
      r.employee?.toLowerCase() === employee.toLowerCase() &&
      r.date   === date &&
      r.status === "Pending"
  );
}

export function isOverlappingLeave(
  employeeName: string,
  newStart:     string,
  newEnd:       string
): { overlaps: boolean; conflictDate?: string } {
  const records = getAllLeaveRequests();
  const active  = records.filter(
    (r) =>
      r.employee?.toLowerCase() === employeeName.toLowerCase() &&
      (r.status === "Pending" || r.status === "Approved")
  );

  const newStartDate = new Date(newStart + "T00:00:00");
  const newEndDate   = new Date((newEnd || newStart) + "T00:00:00");

  for (const record of active) {
    const existStart = new Date(record.date + "T00:00:00");
    const existEnd   = new Date((record.end_date || record.date) + "T00:00:00");
    const overlaps   = newStartDate <= existEnd && newEndDate >= existStart;

    if (overlaps) {
      return {
        overlaps:     true,
        conflictDate: `${record.date}${record.end_date ? " to " + record.end_date : ""}`,
      };
    }
  }

  return { overlaps: false };
}

export function getTodaysAbsences(): LeaveRecord[] {
  const today   = new Date().toISOString().split("T")[0];
  const records = getAllLeaveRequests();
  return records.filter((r) => r.date === today && r.status === "Approved");
}