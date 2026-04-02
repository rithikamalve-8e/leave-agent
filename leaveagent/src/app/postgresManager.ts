import {
  PrismaClient,
  Employee,
  LeaveRequest,
  ConversationRef,
  Holiday,
  AuditLog,
  PendingRequest,
  MonthlySummary,
} from "@prisma/client";

// ── Prisma Client Singleton ────────────────────────────────────────────────

const prisma = new PrismaClient({ log: ["warn", "error"] });
export default prisma;

// ── Types ──────────────────────────────────────────────────────────────────

export interface LeaveBalanceResult {
  requested:  number;
  balance:    number;
  granted:    number;
  lop:        number;
  hasLop:     boolean;
  splits?:    MonthlyLopSplit[];
}

// ADDED
export interface MonthlyLopSplit {
  month:    string;
  days:     number;
  balance:  number;
  lop:      number;
  granted:  number;
}

export interface ConversationRefInput {
  userId:         string;
  userName:       string;
  conversationId: string;
  serviceUrl:     string;
  tenantId?:      string;
  botId:          string;
  // REMOVED: isPersonal
}

export interface PendingRequestInput {
  userId:        string;
  userName:      string;
  intent:        string;
  date:          string;
  end_date?:     string;
  duration:      string;
  days_count:    number;
  lop_days?:     number;   
  reason?:       string;
  balanceResult: LeaveBalanceResult;
  previewCardActivityId?: string | null;
  history:       Array<{ role: "user" | "assistant"; content: string }>;
}

export interface LeaveRequestInput {
  employee:     string;
  email:        string;
  type:         string;
  date:         string;
  end_date?:    string;
  duration:     string;
  days_count:   number;
  lop_days?:    number;    // ADDED
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
    where:   { bot_role: botRole },
    orderBy: { name: "asc" },
  });
}

export async function upsertEmployee(
  data: Omit<Employee, "id" | "created_at" | "updated_at">
): Promise<Employee> {
  return prisma.employee.upsert({
    where:  { name: data.name },
    update: { ...data },
    create: { ...data },
  });
}

export async function adjustLeaveBalance(
  employeeName: string,
  adjustment:   number,
  reason:       string,
  hrName:       string
): Promise<Employee | null> {
  const employee = await findEmployee(employeeName);
  if (!employee) return null;

  const newBalance = Math.max(0, employee.leave_balance + adjustment);
  const updated    = await prisma.employee.update({
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
  console.log(`[DB] Deducted ${days} days from ${employeeName}. Remaining: ${newBalance}`);
}

export async function getUnregisteredEmployees(): Promise<Employee[]> {
  return prisma.employee.findMany({
    where:   { teams_id: null },
    orderBy: { name: "asc" },
  });
}

// ── ADDED: Year start accrual ──────────────────────────────────────────────

export async function runYearStartAccrual(): Promise<void> {
  const now         = new Date();
  const currentYear = now.getFullYear();

  const employees = await getAllEmployees();
  let   count     = 0;

  for (const emp of employees) {
    if (emp.year_entitlement_start === currentYear) continue;

    const prevBalance  = Number(emp.leave_balance ?? 0);
    const carryForward = Math.min(prevBalance, 6);

    await prisma.employee.update({
      where: { name: emp.name },
      data: {
        carry_forward:          carryForward,
        year_entitlement_start: currentYear,
      },
    });
    count++;
    console.log(`[Accrual] ${emp.name}: carry_forward=${carryForward} for ${currentYear}`);
  }

  if (count > 0) console.log(`[Accrual] Year start accrual complete — ${count} employees updated`);
  else           console.log(`[Accrual] Year start already processed`);
}

// ── ADDED: Core balance formula ────────────────────────────────────────────
// balance = carry_forward + (currentMonth × 1.5) - totalUsedThisYear + totalLOPThisYear

export async function getLeaveBalance(
  employeeName: string,
  asOfDate?:    Date
): Promise<number> {
  const now          = asOfDate ?? new Date();
  const currentMonth = now.getMonth() + 1;
  const currentYear  = now.getFullYear();

  const emp = await findEmployee(employeeName);
  if (!emp) return 0;

  const carry = Number(emp.carry_forward ?? 0);

  // All approved non-WFH leaves this year
  const approved = await prisma.leaveRequest.findMany({
    where: {
      employee: { equals: employeeName, mode: "insensitive" },
      status:   "Approved",
      type:     { in: ["LEAVE", "SICK"] },
    },
  });

  const thisYear = approved.filter((r) => {
    const d = new Date(r.date + "T00:00:00");
    return d.getFullYear() === currentYear;
  });

  const totalUsed = thisYear.reduce((sum, r) => sum + (Number(r.days_count) || 0), 0);
  const totalLOP  = thisYear.reduce((sum, r) => sum + (Number(r.lop_days)   || 0), 0);

  const yearEntitlement = currentMonth * 1.5;
  return Math.max(0, carry + yearEntitlement - totalUsed + totalLOP);
}

// ── Leave Requests ─────────────────────────────────────────────────────────

export async function addLeaveRequest(input: LeaveRequestInput): Promise<LeaveRequest> {
  return prisma.leaveRequest.create({
    data: {
      employee:    input.employee,
      email:       input.email,
      type:        input.type,
      date:        input.date,
      end_date:    input.end_date    ?? null,
      duration:    input.duration,
      days_count:  input.days_count,
      lop_days:    input.lop_days   ?? 0,   // ADDED
      reason:      input.reason     ?? null,
      status:      input.status     ?? "Pending",
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
    where:   {
      employee: { equals: employeeName, mode: "insensitive" },
      status:   { not: "Deleted" },
    },
    orderBy: { requested_at: "desc" },
  });
}

export async function getLeaveRequestsForApprover(approverName: string): Promise<LeaveRequest[]> {
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
  employeeName:     string,
  date:             string,
  status:           "Approved" | "Rejected",
  approverName:     string,
  rejectionReason?: string
): Promise<LeaveRequest | null> {
  const record = await prisma.leaveRequest.findFirst({
    where: { employee: { equals: employeeName, mode: "insensitive" }, date, status: "Pending" },
  });
  if (!record) return null;

  const updated = await prisma.leaveRequest.update({
    where: { id: record.id },
    data: {
      status:           status,
      approved_by:      approverName,
      rejection_reason: rejectionReason ?? null,
      lop_days:         status === "Rejected" ? 0 : record.lop_days,
    },
  });

  if (status === "Approved" && ["LEAVE","SICK"].includes(record.type?.toUpperCase())) {
    await deductLeaveBalance(employeeName, record.days_count);
  }

  return updated;
}

export async function deleteLeaveRequest(
  employeeName: string,
  date:         string,
  deletedBy:    string,
  reason:       string
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

export async function isDuplicateRequest(employeeName: string, date: string, duration: string): Promise<boolean> {
    const sameSlot = await prisma.leaveRequest.findFirst({
    where: {
      employee: { equals: employeeName, mode: "insensitive" },
      date,
      duration,
      status: { in: ["Pending", "Approved"] },
    },
  });

  if (sameSlot) return true;

    const fullDay = await prisma.leaveRequest.findFirst({
    where: {
      employee: { equals: employeeName, mode: "insensitive" },
      date,
      duration: "full_day",
      status: { in: ["Pending", "Approved"] },
    },
  });

  if (fullDay) return true;

    if (duration === "full_day") {
    const any = await prisma.leaveRequest.findFirst({
      where: {
        employee: { equals: employeeName, mode: "insensitive" },
        date,
        status: { in: ["Pending", "Approved"] },
      },
    });

    if (any) return true;
  }

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

// ADDED: separate overlap check function
export async function isOverlappingLeave(
  employeeName: string,
  newStart:     string,
  newEnd:       string
): Promise<{ overlaps: boolean; conflictDate?: string }> {
  const active = await prisma.leaveRequest.findMany({
    where: {
      employee: { equals: employeeName, mode: "insensitive" },
      status:   { in: ["Pending", "Approved"] },
    },
  });

  const newStartDate = new Date(newStart + "T00:00:00");
  const newEndDate   = new Date((newEnd || newStart) + "T00:00:00");

  for (const record of active) {
    const existStart = new Date(record.date + "T00:00:00");
    const existEnd   = new Date((record.end_date || record.date) + "T00:00:00");
    if (newStartDate <= existEnd && newEndDate >= existStart) {
      return {
        overlaps:     true,
        conflictDate: `${record.date}${record.end_date ? " to " + record.end_date : ""}`,
      };
    }
  }
  return { overlaps: false };
}

export async function getTodaysAbsences(): Promise<LeaveRequest[]> {
  const today = new Date().toISOString().split("T")[0];
  return prisma.leaveRequest.findMany({
    where:   { date: today, status: "Approved" },
  });
}

export async function getAbsencesForDateRange(
  startDate: string,
  endDate:   string
): Promise<LeaveRequest[]> {
  return prisma.leaveRequest.findMany({
    where: {
      status: "Approved",
      OR: [
        { date: { gte: startDate, lte: endDate }, end_date: null },
        { date: { lte: endDate }, end_date: { gte: startDate } },
      ],
    },
    orderBy: { date: "asc" },
  });
}

export async function getPendingRequestsOlderThanCurrentMonth(): Promise<LeaveRequest[]> {
  const firstOfMonth = new Date(new Date().getFullYear(), new Date().getMonth(), 1);
  return prisma.leaveRequest.findMany({
    where:   { status: "Pending", requested_at: { lt: firstOfMonth } },
    orderBy: { requested_at: "asc" },
  });
}

export async function getMonthlyPendingRequests(): Promise<LeaveRequest[]> {
  const now          = new Date();
  const firstOfMonth = new Date(now.getFullYear(), now.getMonth(), 1);
  return prisma.leaveRequest.findMany({
    where:   { status: "Pending", requested_at: { gte: firstOfMonth, lte: now } },
    orderBy: { requested_at: "asc" },
  });
}

// ── Leave Balance ──────────────────────────────────────────────────────────

// CHANGED: all types except WFH consume balance (was only LEAVE)
export async function checkLeaveBalance(
  employee:      Employee,
  daysRequested: number,
  leaveType:     string,
  startDate?:    string,
  endDate?:      string
): Promise<LeaveBalanceResult> {
  const BALANCE_CONSUMING = [
    "LEAVE","SICK"];

  if (!BALANCE_CONSUMING.includes(leaveType?.toUpperCase())) {
    return {
      requested: daysRequested,
      balance:   999,
      granted:   daysRequested,
      lop:       0,
      hasLop:    false,
    };
  }

  // Cross-month check
  if (startDate && endDate && startDate !== endDate) {
    const startMonth = startDate.substring(0, 7);
    const endMonth   = endDate.substring(0, 7);
    if (startMonth !== endMonth) {
      const result = await calculateCrossMonthLop(employee, startDate, endDate, leaveType);
      return {
        requested: result.totalDays,
        balance:   await getLeaveBalance(employee.name),
        granted:   result.totalGranted,
        lop:       result.totalLop,
        hasLop:    result.totalLop > 0,
        splits:    result.splits,
      };
    }
  }

  // Single month
  const balance = await getLeaveBalance(employee.name);
  const granted = Math.min(daysRequested, balance);
  const lop     = Math.max(0, daysRequested - balance);
  return { requested: daysRequested, balance, granted, lop, hasLop: lop > 0 };
}

// ADDED: Cross-month LOP split
export async function calculateCrossMonthLop(
  employee:  Employee,
  startDate: string,
  endDate:   string,
  leaveType: string
): Promise<{
  splits:       MonthlyLopSplit[];
  totalDays:    number;
  totalLop:     number;
  totalGranted: number;
}> {
  const BALANCE_CONSUMING = [
    "LEAVE","SICK"
  ];
  if (!BALANCE_CONSUMING.includes(leaveType?.toUpperCase())) {
    const days = await countWorkingDays(startDate, endDate);
    return { splits: [], totalDays: days, totalLop: 0, totalGranted: days };
  }

  const start       = new Date(startDate + "T00:00:00");
  const end         = new Date(endDate   + "T00:00:00");
  const current     = new Date(start);
  const monthDays:  Record<string, number> = {};
  const monthOrder: string[] = [];

  while (current <= end) {
    const day = current.getDay();
    if (day !== 0 && day !== 6) {
      const key = `${String(current.getMonth() + 1).padStart(2, "0")}/${current.getFullYear()}`;
      if (!monthDays[key]) { monthDays[key] = 0; monthOrder.push(key); }
      monthDays[key]++;
    }
    current.setDate(current.getDate() + 1);
  }

  const requestDate     = new Date();
  const requestMonthNum = requestDate.getMonth() + 1;
  const requestYear     = requestDate.getFullYear();

  let runningBalance    = await getLeaveBalance(employee.name);
  const splits: MonthlyLopSplit[] = [];

  for (const monthKey of monthOrder) {
    const [mm, yyyy] = monthKey.split("/").map(Number);
    const days       = monthDays[monthKey];

    // Future month → no balance yet (accrual hasn't happened)
    const isFutureMonth =
      yyyy > requestYear ||
      (yyyy === requestYear && mm > requestMonthNum);

    const balanceThisMonth = isFutureMonth ? 0 : runningBalance;
    const granted          = Math.min(days, balanceThisMonth);
    const lop              = Math.max(0, days - balanceThisMonth);

    splits.push({ month: monthKey, days, balance: balanceThisMonth, lop, granted });
    runningBalance = Math.max(0, balanceThisMonth - days);
  }

  const totalDays    = splits.reduce((s, r) => s + r.days,    0);
  const totalLop     = splits.reduce((s, r) => s + r.lop,     0);
  const totalGranted = splits.reduce((s, r) => s + r.granted, 0);

  return { splits, totalDays, totalLop, totalGranted };
}

// ── Working Days ───────────────────────────────────────────────────────────

export async function countWorkingDays(startDate: string, endDate?: string): Promise<number> {
  if (!startDate) return 0;

  const holidays   = await getHolidays();
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
// CHANGED: removed isPersonal

export async function saveConversationRef(input: ConversationRefInput): Promise<void> {
  const existing = await prisma.conversationRef.findUnique({
    where: { userId: input.userId },
  });

  console.log("USER NAME RAW:", JSON.stringify(input.userName));
  const existingIsPersonal = existing?.conversationId?.startsWith("a:") ?? false;
  const newIsPersonal      = input.conversationId.startsWith("a:");

  if (!existing || newIsPersonal || !existingIsPersonal) {
    await prisma.conversationRef.upsert({
      where:  { userId: input.userId },
      update: {
        userName:       input.userName,
        conversationId: input.conversationId,
        serviceUrl:     input.serviceUrl,
        tenantId:       input.tenantId ?? null,
        botId:          input.botId,
      },
      create: {
        userId:         input.userId,
        userName:       input.userName,
        conversationId: input.conversationId,
        serviceUrl:     input.serviceUrl,
        tenantId:       input.tenantId ?? null,
        botId:          input.botId,
      },
    });
    console.log(
      `[DB] Saved conversation ref for ${input.userName} [${newIsPersonal ? "personal" : "group"}]`
    );
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
      end_date:     input.end_date  ?? null,
      duration:     input.duration,
      days_count:   input.days_count,
      lop_days:     input.lop_days  ?? 0,   // ADDED
      reason:       input.reason    ?? null,
      balance_json: JSON.stringify(input.balanceResult),
      history_json: JSON.stringify(input.history),
    },
    create: {
      userId:       input.userId,
      userName:     input.userName,
      intent:       input.intent,
      date:         input.date,
      end_date:     input.end_date  ?? null,
      duration:     input.duration,
      days_count:   input.days_count,
      lop_days:     input.lop_days  ?? 0,   // ADDED
      reason:       input.reason    ?? null,
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
    end_date:      record.end_date    ?? undefined,
    duration:      record.duration,
    days_count:    record.days_count,
    lop_days:      record.lop_days    ?? 0,  // ADDED
    reason:        record.reason      ?? undefined,
    balanceResult: JSON.parse(record.balance_json),
    history:       JSON.parse(record.history_json),
  };
}

export async function getLeaveRequestStatus(
  employeeName: string,
  date: string
): Promise<{ status: "Pending" | "Approved" | "Rejected" } | null> {

  const record = await prisma.leaveRequest.findFirst({
    where: {
      employee: employeeName,
      date: date,
    },
    select: {
      status: true,
    },
  });

  if (!record) return null;

  return {
    status: record.status as "Pending" | "Approved" | "Rejected",
  };
}

export async function clearPendingRequest(userId: string): Promise<void> {
  await prisma.pendingRequest.deleteMany({ where: { userId } });
}

// ── Holidays ───────────────────────────────────────────────────────────────

export async function addHoliday(date: string, name: string, addedBy: string): Promise<Holiday> {
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
  console.log(`[getHolidays] called with month=${month}, year=${year}`);


  // if (month !== undefined || year !== undefined) {
  //   const mm    = String(month).padStart(2, "0");
  //   const yyyy  = String(year);
  //   console.log(`[getHolidays] querying range: ${yyyy}-${mm}-01 → ${yyyy}-${12}-31`);

  //   return prisma.holiday.findMany({
  //     where:   { date: { gte: `${yyyy}-${mm}-01`, lte: `${yyyy}-${12}-31` } },
  //     orderBy: { date: "asc" },
  //   });
  // }
  const today = new Date().toISOString().split("T")[0];
  console.log(`[getHolidays] no month/year — querying from today: ${today}`);

  return prisma.holiday.findMany({
    where:   { date: { gte: today } },
    orderBy: { date: "asc" },
    take:    20,
  });
}

export async function isHoliday(date: string): Promise<boolean> {
  return !!(await prisma.holiday.findUnique({ where: { date } }));
}

export async function clearAllHolidays(clearedBy: string): Promise<number> {
  const result = await prisma.holiday.deleteMany({});

  await appendAuditLog({
    hr_name:         clearedBy,
    action:          "clear_holidays",
    target_employee: null,
    details:         `Cleared all holidays (${result.count} records deleted)`,
  });

  return result.count;
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

// ── Monthly Summary ────────────────────────────────────────────────────────

// ADDED: build one employee's summary row for one month
export async function buildEmployeeMonthlySummary(
  employeeName: string,
  monthNum:     number,
  year:         number
): Promise<{
  month:     string;
  employee:  string;
  opening:   number;
  available: number;
  leaves:    number;
  wfh:       number;
  lop:       number;
  closing:   number;
  pending:   number;
}> {
  const monthKey = `${String(monthNum).padStart(2, "0")}/${year}`;
  const emp      = await findEmployee(employeeName);
  const carry    = Number(emp?.carry_forward ?? 0);

  const allRequests = await prisma.leaveRequest.findMany({
    where: {
      employee: { equals: employeeName, mode: "insensitive" },
      status:   { not: "Deleted" },
    },
  });

  const approvedNonWFH = allRequests.filter((r) => {
    const d = new Date(r.date + "T00:00:00");
    return r.status === "Approved" && r.type !== "WFH" && d.getFullYear() === year;
  });

  const beforeMonth = approvedNonWFH.filter((r) => {
    return new Date(r.date + "T00:00:00").getMonth() + 1 < monthNum;
  });
  const thisMonth = approvedNonWFH.filter((r) => {
    return new Date(r.date + "T00:00:00").getMonth() + 1 === monthNum;
  });

  const totalUsedBeforeM = beforeMonth.reduce((s, r) => s + (Number(r.days_count) || 0), 0);
  const totalLOPBeforeM  = beforeMonth.reduce((s, r) => s + (Number(r.lop_days)   || 0), 0);
  const leavesThisMonth  = thisMonth.reduce((s, r)  => s + (Number(r.days_count) || 0), 0);
  const lopThisMonth     = thisMonth.reduce((s, r)  => s + (Number(r.lop_days)   || 0), 0);

  const wfhThisMonth = allRequests
    .filter((r) => {
      const d = new Date(r.date + "T00:00:00");
      return (
        r.status === "Approved" &&
        r.type   === "WFH" &&
        d.getMonth() + 1 === monthNum &&
        d.getFullYear()   === year
      );
    })
    .reduce((s, r) => s + (Number(r.days_count) || 0), 0);

  const pendingThisMonth = allRequests
    .filter((r) => {
      const d = new Date(r.date + "T00:00:00");
      return (
        r.status === "Pending" &&
        d.getMonth() + 1 === monthNum &&
        d.getFullYear()   === year
      );
    })
    .reduce((s, r) => s + (Number(r.days_count) || 0), 0);

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

// ADDED: build and save for ALL employees for ONE month
export async function buildAndSaveMonthlySummary(
  monthNum: number,
  year:     number
): Promise<void> {
  const employees = await getAllEmployees();
  const monthKey  = `${String(monthNum).padStart(2, "0")}/${year}`;

  for (const emp of employees) {
    const summary = await buildEmployeeMonthlySummary(emp.name, monthNum, year);

    await prisma.monthlySummary.upsert({
      where:  { month_employee: { month: monthKey, employee: emp.name } },
      update: {
        opening:   summary.opening,
        available: summary.available,
        leaves:    summary.leaves,
        wfh:       summary.wfh,
        lop:       summary.lop,
        closing:   summary.closing,
        pending:   summary.pending,
      },
      create: {
        month:     monthKey,
        employee:  emp.name,
        opening:   summary.opening,
        available: summary.available,
        leaves:    summary.leaves,
        wfh:       summary.wfh,
        lop:       summary.lop,
        closing:   summary.closing,
        pending:   summary.pending,
      },
    });
    console.log(`[Summary] ${emp.name} ${monthKey} done`);
  }
  console.log(`[Summary] Built for ${monthKey} — ${employees.length} employees`);
}

// ADDED: for HR report
export async function getMonthlySummaryForMonth(monthKey: string): Promise<MonthlySummary[]> {
  return prisma.monthlySummary.findMany({
    where:   { month: monthKey },
    orderBy: { employee: "asc" },
  });
}

// ── Monthly Report Data ────────────────────────────────────────────────────
// CHANGED: opening balance now uses correct formula

export interface MonthlyReportRow {
  employeeName:     string;
  month:            string;
  openingBalance:   number;
  leaveTaken:       number;
  wfhDays:          number;
  pendingApprovals: number;
  lopDays:          number;   // ADDED
  closingBalance:   number;
}

export async function getMonthlyReportData(
  month: number,
  year:  number
): Promise<MonthlyReportRow[]> {
  const summaries = await getMonthlySummaryForMonth(
    `${String(month).padStart(2, "0")}/${year}`
  );

  return summaries.map((s) => ({
    employeeName:     s.employee,
    month:            s.month,
    openingBalance:   s.opening,   // FIXED: from monthly_summary not reverse-calc
    leaveTaken:       s.leaves,
    wfhDays:          s.wfh,
    pendingApprovals: s.pending,
    lopDays:          s.lop,       // ADDED
    closingBalance:   s.closing,
  }));
}

// ── Graceful Shutdown ──────────────────────────────────────────────────────

export async function disconnectDB(): Promise<void> {
  await prisma.$disconnect();
}

process.on("beforeExit", async () => {
  await disconnectDB();
});

