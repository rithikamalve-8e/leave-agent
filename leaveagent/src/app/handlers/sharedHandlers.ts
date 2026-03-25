import {
  buildHelpCard,
  buildDailySummaryCard,
  buildHolidaysCard,
  buildWhoIsOnLeaveCard,
  formatDisplayDate,
} from "../cards";
import {
  getTodaysAbsences,
  getHolidays,
  getAbsencesForDateRange,
} from "../postgresManager";
import { RoleContext } from "../roleGuard";

export interface CommandContext {
  activity:    any;
  send:        Function;
  api:         any;
  userName:    string;
  userId:      string;
  userMessage: string;
  cmd:         string;
  role:        RoleContext;
}

// ── help ───────────────────────────────────────────────────────────────────

export async function handleHelp(ctx: CommandContext): Promise<void> {
  await ctx.send(buildHelpCard(ctx.role.botRole));
}

// ── summary ────────────────────────────────────────────────────────────────

export async function handleSummary(ctx: CommandContext): Promise<void> {
  const records = await getTodaysAbsences();
  await ctx.send(buildDailySummaryCard(records as any[]));
}

// ── holidays ───────────────────────────────────────────────────────────────

export async function handleHolidays(ctx: CommandContext): Promise<void> {
  const monthNames = [
    "january","february","march","april","may","june",
    "july","august","september","october","november","december",
  ];

  const cmd = ctx.cmd;
  let month: number | undefined;
  let year:  number | undefined;
  let monthLabel: string | undefined;

  for (let i = 0; i < monthNames.length; i++) {
    if (cmd.includes(monthNames[i])) {
      month = i + 1;
      year  = new Date().getFullYear();
      monthLabel = monthNames[i].charAt(0).toUpperCase() + monthNames[i].slice(1) + ` ${year}`;
      break;
    }
  }

  const holidays = await getHolidays(month, year);
  await ctx.send(buildHolidaysCard(holidays as any[], monthLabel));
}

// ── who is on leave today / who is wfh today ──────────────────────────────

export async function handleWhoIsOnLeaveToday(ctx: CommandContext): Promise<void> {
  const today   = new Date().toISOString().split("T")[0];
  const records = await getAbsencesForDateRange(today, today);

  const isWfh = /who is wfh/i.test(ctx.cmd);
  const filtered = isWfh
    ? records.filter((r) => r.type === "WFH")
    : records.filter((r) => r.type !== "WFH");

  const label = isWfh ? "WFH Today" : "On Leave Today";
  const type  = isWfh ? "wfh" : "leave";

  await ctx.send(buildWhoIsOnLeaveCard(filtered as any[], label, type));
}
