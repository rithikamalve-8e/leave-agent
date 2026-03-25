import { RoleContext, getBotRole } from "./roleGuard";
import { CommandContext } from "./handlers/sharedHandlers";
import { NotificationContext } from "./notificationServices";

import {
  handleHelp,
  handleSummary,
  handleHolidays,
  handleWhoIsOnLeaveToday,
} from "./handlers/sharedHandlers";

import {
  handleMyRequests,
  handleMyBalance,
  handleMyStatus,
  handleDeleteRequest,
  handleEditRequest,
} from "./handlers/employeeHandlers";

import {
  handleTeamSummary,
  handleTeamRequests,
  handlePendingApprovals,
  handleWhoIsOnLeave,
  handleWhoIsAvailable,
  handleLeaveHistory,
  handleBalanceForReportee,
  handleApproveLeaveCommand,
  handleRejectLeaveCommand,
} from "./handlers/approveHandlers";

import {
  handleAllRequests,
  handleAllPending,
  handleUnactioned,
  handleApproveAllUnactioned,
  handleApproveUnactionedForEmployee,
  handleRejectAllUnactioned,
  handleOrgSummary,
  handleHRLeaveHistory,
  handleHRBalance,
  handleViewEmployee,
  handleUnregistered,
  handleOrgChart,
  handleTeamOf,
  handleAuditLog,
  handleAdjustBalance,
  handleSetBalance,
  handleResetBalances,
  handleAddLeaveOnBehalf,
  handleHRApproveLeave,
  handleHRRejectLeave,
  handleHRDeleteRequest,
  handleRestoreRequest,
  handleAddHoliday,
  handleEditHoliday,
  handleRescheduleHoliday,
  handleDeleteHoliday,
  handleDownloadReport,
  handleRemindApprovers,
  handleHRWhoIsOnLeaveImpl,
} from "./handlers/hrHandlers";

// ── Command Definition ─────────────────────────────────────────────────────

type BotRole = "employee" | "approver" | "hr";

interface CommandDef {
  match:   (cmd: string, msg: string) => boolean;
  roles:   BotRole[];
  handler: (ctx: CommandContext, nctx: NotificationContext, extras?: any) => Promise<void>;
}

// ── Command Table ──────────────────────────────────────────────────────────

const commands: CommandDef[] = [

  // ── Shared — role-aware inside handler ──────────────────────────────────
  {
    match:   (cmd) => cmd === "help",
    roles:   ["employee", "approver", "hr"],
    handler: handleHelp,
  },
  {
    match:   (cmd) => cmd === "summary",
    roles:   ["employee", "approver", "hr"],
    handler: async (ctx, nctx) => {
      if (ctx.role.isHR)       return handleOrgSummary(ctx);
      if (ctx.role.isApprover) return handleTeamSummary(ctx);
      return handleSummary(ctx);
    },
  },
  {
    match:   (cmd) => cmd === "holidays" || /^holidays\s+\w+/.test(cmd),
    roles:   ["employee", "approver", "hr"],
    handler: handleHolidays,
  },
  {
    match:   (cmd) => /who (is|are) (on leave|absent|wfh|working from home) today/i.test(cmd),
    roles:   ["employee", "approver", "hr"],
    handler: handleWhoIsOnLeaveToday,
  },

  // ── Employee ─────────────────────────────────────────────────────────────
  {
    match:   (cmd) => cmd === "my requests" || cmd === "my requests all",
    roles:   ["employee", "approver", "hr"],
    handler: handleMyRequests,
  },
  {
    match:   (cmd) => /my balance|leave balance|my leave balance|^balance$|how many (leave|days)|leaves left|days (remaining|left)/i.test(cmd),
    roles:   ["employee", "approver", "hr"],
    handler: handleMyBalance,
  },
  {
    match:   (cmd) => /^my status/i.test(cmd),
    roles:   ["employee", "approver", "hr"],
    handler: handleMyStatus,
  },
  {
    match:   (cmd, msg) => /^cancel request|^delete request/i.test(cmd) && !/ \w+ \d{4}-\d{2}-\d{2}/.test(msg),
    roles:   ["employee", "approver"],
    handler: handleDeleteRequest,
  },
  {
    match:   (cmd) => /^edit request/i.test(cmd),
    roles:   ["employee", "approver", "hr"],
    handler: (ctx, nctx, extras) => handleEditRequest(ctx, extras?.savePendingRequest, extras?.getPendingRequest),
  },

  // ── Approver ─────────────────────────────────────────────────────────────
  {
    match:   (cmd) => /^team summary/i.test(cmd),
    roles:   ["approver", "hr"],
    handler: handleTeamSummary,
  },
  {
    match:   (cmd) => /^team requests/i.test(cmd),
    roles:   ["approver", "hr"],
    handler: handleTeamRequests,
  },
  {
    match:   (cmd) => cmd === "pending approvals",
    roles:   ["approver", "hr"],
    handler: handlePendingApprovals,
  },
  {
    match:   (cmd) => /who (is|are) (on leave|absent|wfh|working from home)/i.test(cmd) && !/today/i.test(cmd),
    roles:   ["approver", "hr"],
    handler: async (ctx, nctx) => {
      if (ctx.role.isHR) return handleHRWhoIsOnLeaveImpl(ctx);
      return handleWhoIsOnLeave(ctx, true);
    },
  },
  {
    match:   (cmd) => /who is available/i.test(cmd),
    roles:   ["approver", "hr"],
    handler: handleWhoIsAvailable,
  },
  {
    match:   (cmd, msg) => /^leave history/i.test(cmd),
    roles:   ["approver", "hr"],
    handler: async (ctx, nctx) => {
      if (ctx.role.isHR) return handleHRLeaveHistory(ctx);
      return handleLeaveHistory(ctx);
    },
  },
  {
    match:   (cmd) => /^balance \w/i.test(cmd),
    roles:   ["approver", "hr"],
    handler: async (ctx, nctx) => {
      if (ctx.role.isHR) return handleHRBalance(ctx);
      return handleBalanceForReportee(ctx);
    },
  },
  {
    match:   (cmd, msg) => /^approve leave\s+/i.test(msg) && !/all unactioned/i.test(msg),
    roles:   ["approver", "hr"],
    handler: async (ctx, nctx) => {
      if (ctx.role.isHR) return handleHRApproveLeave(ctx, nctx);
      return handleApproveLeaveCommand(ctx, nctx);
    },
  },
  {
    match:   (cmd, msg) => /^reject leave\s+/i.test(msg) && !/all unactioned/i.test(msg),
    roles:   ["approver", "hr"],
    handler: async (ctx, nctx) => {
      if (ctx.role.isHR) return handleHRRejectLeave(ctx, nctx);
      return handleRejectLeaveCommand(ctx, nctx);
    },
  },

  // ── HR only ──────────────────────────────────────────────────────────────
  {
    match:   (cmd) => cmd === "all requests" || /^all requests/i.test(cmd),
    roles:   ["hr"],
    handler: handleAllRequests,
  },
  {
    match:   (cmd) => cmd === "pending" || /^pending\s+\w/i.test(cmd),
    roles:   ["hr"],
    handler: handleAllPending,
  },
  {
    match:   (cmd) => cmd === "unactioned",
    roles:   ["hr"],
    handler: handleUnactioned,
  },
  {
    match:   (cmd) => cmd === "approve all unactioned",
    roles:   ["hr"],
    handler: handleApproveAllUnactioned,
  },
  {
    match:   (cmd) => /^approve unactioned\s+/i.test(cmd),
    roles:   ["hr"],
    handler: handleApproveUnactionedForEmployee,
  },
  {
    match:   (cmd) => /^reject all unactioned/i.test(cmd),
    roles:   ["hr"],
    handler: handleRejectAllUnactioned,
  },
  {
    match:   (cmd) => cmd === "org summary" || /^org summary/i.test(cmd),
    roles:   ["hr"],
    handler: handleOrgSummary,
  },
  {
    match:   (cmd) => /^view employee/i.test(cmd),
    roles:   ["hr"],
    handler: handleViewEmployee,
  },
  {
    match:   (cmd) => cmd === "unregistered",
    roles:   ["hr"],
    handler: handleUnregistered,
  },
  {
    match:   (cmd) => cmd === "org chart",
    roles:   ["hr"],
    handler: handleOrgChart,
  },
  {
    match:   (cmd) => /^team\s+\w/i.test(cmd) && !/^team (summary|requests)/i.test(cmd),
    roles:   ["hr"],
    handler: handleTeamOf,
  },
  {
    match:   (cmd) => /^audit log/i.test(cmd),
    roles:   ["hr"],
    handler: handleAuditLog,
  },
  {
    match:   (cmd) => /^adjust balance/i.test(cmd),
    roles:   ["hr"],
    handler: handleAdjustBalance,
  },
  {
    match:   (cmd) => /^set balance/i.test(cmd),
    roles:   ["hr"],
    handler: handleSetBalance,
  },
  {
    match:   (cmd) => /^reset balances/i.test(cmd),
    roles:   ["hr"],
    handler: handleResetBalances,
  },
  {
    match:   (cmd) => /^add leave for/i.test(cmd),
    roles:   ["hr"],
    handler: handleAddLeaveOnBehalf,
  },
  {
    match:   (cmd, msg) => /^delete request\s+.+\s+\d{4}-\d{2}-\d{2}/i.test(msg),
    roles:   ["hr"],
    handler: handleHRDeleteRequest,
  },
  {
    match:   (cmd) => /^restore request/i.test(cmd),
    roles:   ["hr"],
    handler: handleRestoreRequest,
  },
  {
    match:   (cmd) => /^add holiday/i.test(cmd),
    roles:   ["hr"],
    handler: handleAddHoliday,
  },
  {
    match:   (cmd) => /^edit holiday/i.test(cmd),
    roles:   ["hr"],
    handler: handleEditHoliday,
  },
  {
    match:   (cmd) => /^reschedule holiday/i.test(cmd),
    roles:   ["hr"],
    handler: handleRescheduleHoliday,
  },
  {
    match:   (cmd) => /^delete holiday/i.test(cmd),
    roles:   ["hr"],
    handler: handleDeleteHoliday,
  },
  {
    match:   (cmd) => /^download report|^report\s+/i.test(cmd),
    roles:   ["hr"],
    handler: handleDownloadReport,
  },
  {
    match:   (cmd) => cmd === "remind approvers",
    roles:   ["hr"],
    handler: handleRemindApprovers,
  },
];

// ── Router ─────────────────────────────────────────────────────────────────

export async function routeCommand(
  ctx:    CommandContext,
  nctx:   NotificationContext,
  extras?: any
): Promise<boolean> {
  for (const def of commands) {
    if (!def.match(ctx.cmd, ctx.userMessage)) continue;

    // Role check — HR always passes
    if (ctx.role.botRole !== "hr" && !def.roles.includes(ctx.role.botRole)) {
      await ctx.send(`You don't have permission to use that command. Type \`help\` to see your available commands.`);
      return true;
    }

    await def.handler(ctx, nctx, extras);
    return true;
  }

  return false; // no command matched — fall through to AI intent parser
}
