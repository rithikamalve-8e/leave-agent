import { PrismaClient, Employee, LeaveRequest, ConversationRef, Holiday, AuditLog, PendingRequest } from "@prisma/client";

// ── Prisma Client Singleton ────────────────────────────────────────────────

const prisma = new PrismaClient({
  log: ["warn", "error"],
});

export default prisma;

// ── Types ──────────────────────────────────────────────────────────────────

export interface LeaveBalanceResult {
  requested: number;
  balance:   number;
  granted:   number;
  lop:       number;
  hasLop:    boolean;
}

export interface ConversationRefInput {
  userId:         string;
  userName:       string;
  conversationId: string;
  serviceUrl:     string;
  tenantId?:      string;
  botId:          string;
  isPersonal?:    boolean;
}

export interface PendingRequestInput {
  userId:       string;
  userName:     string;
  intent:       string;
  date:         string;
  end_date?:    string;
  duration:     string;
  days_count:   number;
  reason?:      string;
  balanceResult: LeaveBalanceResult;
  history:      Array<{ role: "user" | "assistant"; content: string }>;
}

export interface LeaveRequestInput {
  employee:     string;
  email:        string;
  type:         string;
  date:         string;
  end_date?:    string;
  duration:     string;
  days_count:   number;
  reason?:      string;
  status?:      string;
  approved_by?: string;
}

// ── Employee ───────────────────────────────────────────────────────────────

export async function findEmployee(name: string): Promise<Employee | null> {
  return prisma.employee.findFirst({
    where: { name: { equals: name, mode: "insensitive" } },
  });
}

export async function findEmployeeByTeamsId(teamsId: string): Promise<Employee | null> {
  return prisma.employee.findFirst({
    where: { teams_id: teamsId },
  });
}

export async function getAllEmployees(): Promise<Employee[]> {
  return prisma.employee.findMany({ orderBy: { name: "asc" } });
}

export async function getEmployeesByBotRole(botRole: string): Promise<Employee[]> {
  return prisma.employee.findMany({
    where: { bot_role: botRole },
    orderBy: { name: "asc" },
  });
}

export async function upsertEmployee(data: Omit<Employee, "id" | "created_at" | "updated_at">): Promise<Employee> {
  return prisma.employee.upsert({
    where: { name: data.name },
    update: { ...data },
    create: { ...data },
  });
}

export async function adjustLeaveBalance(
  employeeName: string,
  adjustment: number,
  reason: string,
  hrName: string
): Promise<Employee | null> {
  const employee = await findEmployee(employeeName);
  if (!employee) return null;

  const newBalance = Math.max(0, employee.leave_balance + adjustment);

  const updated = await prisma.employee.update({
    where: { name: employee.name },
    data:  { leave_balance: newBalance },
  });

  await appendAuditLog({
    hr_name:         hrName,
    action:          "balance_adjust",
    target_employee: employeeName,
    details:         `${adjustment > 0 ? "+" : ""}${adjustment} days. Reason: ${reason}. New balance: ${newBalance}`,
  });

  return updated;
}

export async function deductLeaveBalance(employeeName: string, days: number): Promise<void> {
  const employee = await findEmployee(employeeName);
  if (!employee) return;

  const newBalance = Math.max(0, employee.leave_balance - days);
  await prisma.employee.update({
    where: { name: employee.name },
    data:  { leave_balance: newBalance },
  });

  console.log(`[DB] Deducted ${days} leave days from ${employeeName}. Remaining: ${newBalance}`);
}

export async function getUnregisteredEmployees(): Promise<Employee[]> {
  return prisma.employee.findMany({
    where: { teams_id: null },
    orderBy: { name: "asc" },
  });
}

// ── Leave Requests ─────────────────────────────────────────────────────────

export async function addLeaveRequest(input: LeaveRequestInput): Promise<LeaveRequest> {
  const employee = await findEmployee(input.employee);
  return prisma.leaveRequest.create({
    data: {
      employee:   input.employee,
      email:      input.email,
      type:       input.type,
      date:       input.date,
      end_date:   input.end_date ?? null,
      duration:   input.duration,
      days_count: input.days_count,
      reason:     input.reason ?? null,
      status:     input.status ?? "Pending",
      approved_by: input.approved_by ?? null,
    },
  });
}

export async function getAllLeaveRequests(): Promise<LeaveRequest[]> {
  return prisma.leaveRequest.findMany({
    where:   { status: { not: "Deleted" } },
    orderBy: { requested_at: "desc" },
  });
}

export async function getLeaveRequestsByEmployee(employeeName: string): Promise<LeaveRequest[]> {
  return prisma.leaveRequest.findMany({
    where:   { employee: { equals: employeeName, mode: "insensitive" }, status: { not: "Deleted" } },
    orderBy: { requested_at: "desc" },
  });
}

export async function getLeaveRequestsForApprover(approverName: string): Promise<LeaveRequest[]> {
  // Get all employees who report to this approver
  const reportees = await prisma.employee.findMany({
    where: {
      OR: [
        { teamlead:  { equals: approverName, mode: "insensitive" } },
        { manager:   { equals: approverName, mode: "insensitive" } },
      ],
    },
  });

  const names = reportees.map((e) => e.name);
  return prisma.leaveRequest.findMany({
    where:   { employee: { in: names }, status: { not: "Deleted" } },
    orderBy: { requested_at: "desc" },
  });
}

export async function getPendingRequestsForApprover(approverName: string): Promise<LeaveRequest[]> {
  const reportees = await prisma.employee.findMany({
    where: {
      OR: [
        { teamlead: { equals: approverName, mode: "insensitive" } },
        { manager:  { equals: approverName, mode: "insensitive" } },
      ],
    },
  });

  const names = reportees.map((e) => e.name);
  return prisma.leaveRequest.findMany({
    where:   { employee: { in: names }, status: "Pending" },
    orderBy: { requested_at: "asc" },
  });
}

export async function updateLeaveStatus(
  employeeName: string,
  date: string,
  status: "Approved" | "Rejected",
  approverName: string,
  rejectionReason?: string
): Promise<LeaveRequest | null> {
  const record = await prisma.leaveRequest.findFirst({
    where: {
      employee: { equals: employeeName, mode: "insensitive" },
      date,
      status:   "Pending",
    },
  });

  if (!record) return null;

  const updated = await prisma.leaveRequest.update({
    where: { id: record.id },
    data: {
      status:           status,
      approved_by:      approverName,
      rejection_reason: rejectionReason ?? null,
    },
  });

  // Deduct balance on approval for LEAVE type
  if (status === "Approved" && record.type === "LEAVE") {
    await deductLeaveBalance(employeeName, record.days_count);
  }

  return updated;
}

export async function deleteLeaveRequest(
  employeeName: string,
  date: string,
  deletedBy: string,
  reason: string
): Promise<LeaveRequest | null> {
  const record = await prisma.leaveRequest.findFirst({
    where: {
      employee: { equals: employeeName, mode: "insensitive" },
      date,
      status:   { not: "Deleted" },
    },
  });

  if (!record) return null;

  const updated = await prisma.leaveRequest.update({
    where: { id: record.id },
    data: {
      status:     "Deleted",
      deleted_by: deletedBy,
      deleted_at: new Date(),
    },
  });

  await appendAuditLog({
    hr_name:         deletedBy,
    action:          "delete_request",
    target_employee: employeeName,
    details:         `Deleted ${record.type} on ${date}. Reason: ${reason}`,
  });

  return updated;
}

export async function isDuplicateRequest(employeeName: string, date: string): Promise<boolean> {
  // Check exact date match
  const exact = await prisma.leaveRequest.findFirst({
    where: {
      employee: { equals: employeeName, mode: "insensitive" },
      date,
      status:   { in: ["Pending", "Approved"] },
    },
  });
  if (exact) return true;

  // Check if date falls within an existing multi-day request window
  const multiDay = await prisma.leaveRequest.findMany({
    where: {
      employee: { equals: employeeName, mode: "insensitive" },
      status:   { in: ["Pending", "Approved"] },
      end_date: { not: null },
    },
  });

  for (const r of multiDay) {
    if (r.end_date && r.date <= date && date <= r.end_date) return true;
  }

  return false;
}

export async function getTodaysAbsences(): Promise<LeaveRequest[]> {
  const today = new Date().toISOString().split("T")[0];
  return prisma.leaveRequest.findMany({
    where: {
      date:   today,
      status: "Approved",
    },
  });
}

export async function getAbsencesForDateRange(
  startDate: string,
  endDate: string
): Promise<LeaveRequest[]> {
  return prisma.leaveRequest.findMany({
    where: {
      status: "Approved",
      OR: [
        // Single day requests within range
        { date: { gte: startDate, lte: endDate }, end_date: null },
        // Multi-day requests overlapping with range
        { date: { lte: endDate }, end_date: { gte: startDate } },
      ],
    },
    orderBy: { date: "asc" },
  });
}

export async function getPendingRequestsOlderThanCurrentMonth(): Promise<LeaveRequest[]> {
  const now        = new Date();
  const firstOfMonth = new Date(now.getFullYear(), now.getMonth(), 1);

  return prisma.leaveRequest.findMany({
    where: {
      status:       "Pending",
      requested_at: { lt: firstOfMonth },
    },
    orderBy: { requested_at: "asc" },
  });
}

export async function getMonthlyPendingRequests(): Promise<LeaveRequest[]> {
  const now          = new Date();
  const firstOfMonth = new Date(now.getFullYear(), now.getMonth(), 1);
  const today        = new Date();

  return prisma.leaveRequest.findMany({
    where: {
      status:       "Pending",
      requested_at: { gte: firstOfMonth, lte: today },
    },
    orderBy: { requested_at: "asc" },
  });
}

// ── Leave Balance ──────────────────────────────────────────────────────────

export function checkLeaveBalance(
  employee: Employee,
  daysRequested: number,
  leaveType: string
): LeaveBalanceResult {
  if (leaveType !== "LEAVE") {
    return { requested: daysRequested, balance: 999, granted: daysRequested, lop: 0, hasLop: false };
  }

  const balance = Number(employee.leave_balance ?? 0);
  const granted = Math.min(daysRequested, balance);
  const lop     = Math.max(0, daysRequested - balance);

  return { requested: daysRequested, balance, granted, lop, hasLop: lop > 0 };
}

// ── Working Days ───────────────────────────────────────────────────────────

export async function countWorkingDays(startDate: string, endDate?: string): Promise<number> {
  if (!startDate) return 0;

  const holidays = await getHolidays();
  const holidaySet = new Set(holidays.map((h) => h.date));

  const start   = new Date(startDate + "T00:00:00");
  const end     = endDate ? new Date(endDate + "T00:00:00") : new Date(startDate + "T00:00:00");
  let   count   = 0;
  const current = new Date(start);

  while (current <= end) {
    const day     = current.getDay();
    const dateStr = current.toISOString().split("T")[0];
    if (day !== 0 && day !== 6 && !holidaySet.has(dateStr)) count++;
    current.setDate(current.getDate() + 1);
  }

  return count;
}

// ── Conversation Refs ──────────────────────────────────────────────────────

export async function saveConversationRef(input: ConversationRefInput): Promise<void> {
  const existing = await prisma.conversationRef.findUnique({
    where: { userId: input.userId },
  });

  const existingIsPersonal = existing?.isPersonal ?? false;
  const newIsPersonal      = input.isPersonal ?? false;

  // Only overwrite if: no existing, OR new is personal, OR existing is group
  if (!existing || newIsPersonal || !existingIsPersonal) {
    await prisma.conversationRef.upsert({
      where:  { userId: input.userId },
      update: { ...input, isPersonal: newIsPersonal },
      create: { ...input, isPersonal: newIsPersonal },
    });
    console.log(`[DB] Saved conversation ref for ${input.userName} [${newIsPersonal ? "personal" : "group"}]`);
  }
}

export async function getConversationRef(userId: string): Promise<ConversationRef | null> {
  return prisma.conversationRef.findUnique({ where: { userId } });
}

export async function getAllConversationRefs(): Promise<ConversationRef[]> {
  return prisma.conversationRef.findMany();
}

// ── Pending Requests ───────────────────────────────────────────────────────

export async function savePendingRequest(input: PendingRequestInput): Promise<void> {
  await prisma.pendingRequest.upsert({
    where:  { userId: input.userId },
    update: {
      intent:       input.intent,
      date:         input.date,
      end_date:     input.end_date ?? null,
      duration:     input.duration,
      days_count:   input.days_count,
      reason:       input.reason ?? null,
      balance_json: JSON.stringify(input.balanceResult),
      history_json: JSON.stringify(input.history),
    },
    create: {
      userId:       input.userId,
      userName:     input.userName,
      intent:       input.intent,
      date:         input.date,
      end_date:     input.end_date ?? null,
      duration:     input.duration,
      days_count:   input.days_count,
      reason:       input.reason ?? null,
      balance_json: JSON.stringify(input.balanceResult),
      history_json: JSON.stringify(input.history),
    },
  });
}

export async function getPendingRequest(userId: string): Promise<PendingRequestInput | null> {
  const record = await prisma.pendingRequest.findUnique({ where: { userId } });
  if (!record) return null;

  return {
    userId:        record.userId,
    userName:      record.userName,
    intent:        record.intent,
    date:          record.date,
    end_date:      record.end_date ?? undefined,
    duration:      record.duration,
    days_count:    record.days_count,
    reason:        record.reason ?? undefined,
    balanceResult: JSON.parse(record.balance_json),
    history:       JSON.parse(record.history_json),
  };
}

export async function clearPendingRequest(userId: string): Promise<void> {
  await prisma.pendingRequest.deleteMany({ where: { userId } });
}

// ── Holidays ───────────────────────────────────────────────────────────────

export async function addHoliday(
  date: string,
  name: string,
  addedBy: string
): Promise<Holiday> {
  const holiday = await prisma.holiday.upsert({
    where:  { date },
    update: { name, added_by: addedBy },
    create: { date, name, added_by: addedBy },
  });

  await appendAuditLog({
    hr_name:         addedBy,
    action:          "add_holiday",
    target_employee: null,
    details:         `Added holiday: ${name} on ${date}`,
  });

  return holiday;
}

export async function getHolidays(month?: number, year?: number): Promise<Holiday[]> {
  if (month !== undefined && year !== undefined) {
    const mm    = String(month).padStart(2, "0");
    const yyyy  = String(year);
    const start = `${yyyy}-${mm}-01`;
    const end   = `${yyyy}-${mm}-31`;
    return prisma.holiday.findMany({
      where:   { date: { gte: start, lte: end } },
      orderBy: { date: "asc" },
    });
  }

  // Default: upcoming holidays from today
  const today = new Date().toISOString().split("T")[0];
  return prisma.holiday.findMany({
    where:   { date: { gte: today } },
    orderBy: { date: "asc" },
    take:    20,
  });
}

export async function isHoliday(date: string): Promise<boolean> {
  const holiday = await prisma.holiday.findUnique({ where: { date } });
  return !!holiday;
}

// ── Audit Log ──────────────────────────────────────────────────────────────

export async function appendAuditLog(entry: {
  hr_name:          string;
  action:           string;
  target_employee?: string | null;
  details:          string;
}): Promise<void> {
  await prisma.auditLog.create({
    data: {
      hr_name:         entry.hr_name,
      action:          entry.action,
      target_employee: entry.target_employee ?? null,
      details:         entry.details,
    },
  });
}

export async function getAuditLog(limit = 50): Promise<AuditLog[]> {
  return prisma.auditLog.findMany({
    orderBy: { timestamp: "desc" },
    take:    limit,
  });
}

// ── Monthly Report Data ────────────────────────────────────────────────────

export interface MonthlyReportRow {
  employeeName:      string;
  month:             string;
  openingBalance:    number;
  leaveTaken:        number;
  wfhDays:           number;
  pendingApprovals:  number;
  closingBalance:    number;
}

export async function getMonthlyReportData(
  month: number,
  year: number
): Promise<MonthlyReportRow[]> {
  const mm    = String(month).padStart(2, "0");
  const yyyy  = String(year);
  const start = `${yyyy}-${mm}-01`;
  const end   = `${yyyy}-${mm}-31`;

  const employees = await getAllEmployees();
  const requests  = await prisma.leaveRequest.findMany({
    where: { date: { gte: start, lte: end } },
  });

  const rows: MonthlyReportRow[] = [];

  for (const emp of employees) {
    const empReqs = requests.filter(
      (r) => r.employee.toLowerCase() === emp.name.toLowerCase()
    );

    const leaveTaken = empReqs
      .filter((r) => r.type !== "WFH" && r.status === "Approved")
      .reduce((sum, r) => sum + r.days_count, 0);

    const wfhDays = empReqs
      .filter((r) => r.type === "WFH" && r.status === "Approved")
      .reduce((sum, r) => sum + r.days_count, 0);

    const pendingApprovals = empReqs
      .filter((r) => r.status === "Pending")
      .reduce((sum, r) => sum + r.days_count, 0);

    const openingBalance  = emp.leave_balance + leaveTaken; // reverse-calculate
    const closingBalance  = emp.leave_balance;

    rows.push({
      employeeName:     emp.name,
      month:            `${yyyy}-${mm}`,
      openingBalance,
      leaveTaken,
      wfhDays,
      pendingApprovals,
      closingBalance,
    });
  }

  return rows;
}

// ── Graceful Shutdown ──────────────────────────────────────────────────────

export async function disconnectDB(): Promise<void> {
  await prisma.$disconnect();
}

process.on("beforeExit", async () => {
  await disconnectDB();
});