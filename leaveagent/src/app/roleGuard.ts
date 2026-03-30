import { findEmployee } from "./postgresManager";

// ── Types ──────────────────────────────────────────────────────────────────

export type BotRole = "employee" | "approver" | "hr";

export interface RoleContext {
  userName:   string;
  botRole:    BotRole;
  isHR:       boolean;
  isApprover: boolean;
  isEmployee: boolean;
}

// ── Role Detection ─────────────────────────────────────────────────────────

export async function getBotRole(userName: string): Promise<BotRole> {
  const employee = await findEmployee(userName);
  if (!employee) return "employee";

  const raw = (employee.bot_role ?? "").toString().toLowerCase().trim();

  if (raw === "hr")       return "hr";
  if (raw === "approver") return "approver";
  return "employee";
}

export async function getRoleContext(userName: string): Promise<RoleContext> {
  const botRole = await getBotRole(userName);
  return {
    userName,
    botRole,
    isHR:       botRole === "hr",
    isApprover: botRole === "approver" || botRole === "hr",
    isEmployee: true,
  };
}

// ── Permission Guards ──────────────────────────────────────────────────────

export async function hasRole(userName: string, ...requiredRoles: BotRole[]): Promise<boolean> {
  const role = await getBotRole(userName);
  if (role === "hr") return true;
  return requiredRoles.includes(role);
}

export async function isHR(userName: string): Promise<boolean> {
  return (await getBotRole(userName)) === "hr";
}

export async function isApprover(userName: string): Promise<boolean> {
  const role = await getBotRole(userName);
  return role === "approver" || role === "hr";
}

export async function canApprove(approverName: string, employeeName: string): Promise<boolean> {
  if (approverName.toLowerCase() === employeeName.toLowerCase()) return false;

  const role = await getBotRole(approverName);

  if (role === "hr") return true;

  if (role === "approver") {
    const employee = await findEmployee(employeeName);
    if (!employee) return false;

    const teamlead = employee.teamlead?.toLowerCase();
    const manager  = employee.manager?.toLowerCase();
    const name     = approverName.toLowerCase();

    return teamlead === name || manager === name;
  }

  return false;
}

export async function canDelete(
  userName:      string,
  employeeName:  string,
  requestStatus: string
): Promise<boolean> {
  const role = await getBotRole(userName);

  if (role === "hr") return true;

  if (role === "employee" || role === "approver") {
    const isOwn     = userName.toLowerCase() === employeeName.toLowerCase();
    const isPending = requestStatus?.toLowerCase() === "pending";
    return isOwn && isPending;
  }

  return false;
}

export async function canEdit(
  userName:      string,
  employeeName:  string,
  requestStatus: string
): Promise<boolean> {
  return canDelete(userName, employeeName, requestStatus);
}

export async function getRoleLabel(userName: string): Promise<string> {
  const role = await getBotRole(userName);
  if (role === "hr")       return "HR";
  if (role === "approver") return "Approver";
  return "Employee";
}

// ── Policy Bypass ──────────────────────────────────────────────────────────

export async function bypassesPolicy(userName: string): Promise<boolean> {
  return isHR(userName);
}