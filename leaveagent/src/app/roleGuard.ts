import { findEmployee } from "./excelManager";

// ── Types ──────────────────────────────────────────────────────────────────

export type BotRole = "employee" | "approver" | "hr";

export interface RoleContext {
  userName:  string;
  botRole:   BotRole;
  isHR:      boolean;
  isApprover: boolean;
  isEmployee: boolean;
}

// ── Role Detection ─────────────────────────────────────────────────────────

/**
 * Returns the bot_role for a given user name.
 * Falls back to "employee" if not found or bot_role not set.
 */
export function getBotRole(userName: string): BotRole {
  const employee = findEmployee(userName);
  if (!employee) return "employee";

  const raw = ((employee as any).bot_role ?? "").toString().toLowerCase().trim();

  if (raw === "hr")       return "hr";
  if (raw === "approver") return "approver";
  return "employee";
}

/**
 * Returns a full RoleContext object for a user — use this in handlers
 * to avoid calling getBotRole multiple times.
 */
export function getRoleContext(userName: string): RoleContext {
  const botRole = getBotRole(userName);
  return {
    userName,
    botRole,
    isHR:       botRole === "hr",
    isApprover: botRole === "approver" || botRole === "hr", // HR can do everything approver can
    isEmployee: true, // everyone can do employee actions
  };
}

// ── Permission Guards ──────────────────────────────────────────────────────

/**
 * Returns true if the user has at least one of the required roles.
 * HR always passes — they have access to everything.
 */
export function hasRole(userName: string, ...requiredRoles: BotRole[]): boolean {
  const role = getBotRole(userName);
  if (role === "hr") return true; // HR bypasses all role checks
  return requiredRoles.includes(role);
}

/**
 * Returns true if the user is HR.
 */
export function isHR(userName: string): boolean {
  return getBotRole(userName) === "hr";
}

/**
 * Returns true if the user is an approver or HR.
 */
export function isApprover(userName: string): boolean {
  const role = getBotRole(userName);
  return role === "approver" || role === "hr";
}

/**
 * Returns true if the user can approve a specific employee's request.
 * Approvers can only approve their direct reportees.
 * HR can approve anyone.
 * Nobody can approve their own request.
 */
export function canApprove(approverName: string, employeeName: string): boolean {
  // Nobody approves their own request
  if (approverName.toLowerCase() === employeeName.toLowerCase()) return false;

  const role = getBotRole(approverName);

  // HR can approve anyone
  if (role === "hr") return true;

  // Approver can only approve their direct reportees
  if (role === "approver") {
    const employee = findEmployee(employeeName);
    if (!employee) return false;

    const teamlead = employee.teamlead?.toLowerCase();
    const manager  = employee.manager?.toLowerCase();
    const name     = approverName.toLowerCase();

    return teamlead === name || manager === name;
  }

  return false;
}

/**
 * Returns true if the user can delete a specific leave request.
 * - Employee: only their own request, only while Pending
 * - Approver: cannot delete
 * - HR: can delete any request regardless of status
 */
export function canDelete(
  userName: string,
  employeeName: string,
  requestStatus: string
): boolean {
  const role = getBotRole(userName);

  if (role === "hr") return true;

  if (role === "employee" || role === "approver") {
    const isOwn    = userName.toLowerCase() === employeeName.toLowerCase();
    const isPending = requestStatus?.toLowerCase() === "pending";
    return isOwn && isPending;
  }

  return false;
}

/**
 * Returns true if the user can edit a specific leave request.
 * Same rules as delete — only owner while Pending, or HR anytime.
 */
export function canEdit(
  userName: string,
  employeeName: string,
  requestStatus: string
): boolean {
  return canDelete(userName, employeeName, requestStatus);
}

/**
 * Returns a user-friendly role label for display.
 */
export function getRoleLabel(userName: string): string {
  const role = getBotRole(userName);
  if (role === "hr")       return "HR";
  if (role === "approver") return "Approver";
  return "Employee";
}

// ── Policy Bypass ──────────────────────────────────────────────────────────

/**
 * Returns true if the user should bypass policy checks.
 * HR bypasses: past dates, 42-day marriage rule, weekend blocks, etc.
 */
export function bypassesPolicy(userName: string): boolean {
  return isHR(userName);
}
