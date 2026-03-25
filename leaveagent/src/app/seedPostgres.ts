/**
 * scripts/seedPostgres.ts
 *
 * Seeds the Neon Postgres database with:
 * - Employees (from hardcoded list — update with real data)
 * - 2026 company holidays
 *
 * Run: npx ts-node scripts/seedPostgres.ts
 */

import * as dotenv from "dotenv";
import * as path   from "path";
dotenv.config({ path: path.resolve(process.cwd(), ".env") });
dotenv.config({ path: path.resolve(process.cwd(), "env/.env.local") });

import { PrismaClient } from "@prisma/client";

const prisma = new PrismaClient();

// ── Employees ──────────────────────────────────────────────────────────────
// Update teams_id values after each person messages the bot once.
// Terminal will print: [LeaveAgent] "hi" from Rithika MR (29:1abc...)

const employees = [
  {
    name:               "Rithika MR",
    email:              "rithika.mr@8thelement.ai",
    role:               "employee",
    bot_role:           "employee",
    manager:            "Varsha M",
    manager_email:      "varsha.m@8thelement.ai",
    manager_teams_id:   "29:1PXmmVnyMqR2UhE7ZjTilKBqauLAV9GJMGcVBN_o2OKMlNe605D1hkuPVUNJPAxRmy1ZEs6I7uv3uV5BqMJ2GSA",
    teamlead:           "Varsha M",
    teamlead_email:     "varsha.m@8thelement.ai",
    teamlead_teams_id:  "29:1PXmmVnyMqR2UhE7ZjTilKBqauLAV9GJMGcVBN_o2OKMlNe605D1hkuPVUNJPAxRmy1ZEs6I7uv3uV5BqMJ2GSA",
    teams_id:           "29:1zhczMgvxoe76A_tv8PSffaRP1L-AcbVHFfYjdSvVr2u9ooGzSUIJejBSOaEpHpo6Dyw8ZNGeTj8GdVcr3c2dqQ",
    leave_balance:      19,
    carry_forward:      0,
  },
  {
    name:               "Varsha M",
    email:              "varsha.m@8thelement.ai",
    role:               "teamlead",
    bot_role:           "approver",
    manager:            "Tushar Thapliyal",
    manager_email:      "tushar@8thelement.ai",
    manager_teams_id:   "29:1gCB8FOCd4_AdYy8lCvQBty0jW1Cs6xN1wnU4wMs2xJAQjesC2VhTR1PIt_Luo-yBLmHQ_PjAb0xg2ETM0mc-kg",
    teamlead:           "",
    teamlead_email:     "",
    teamlead_teams_id:  "",
    teams_id:           "29:1PXmmVnyMqR2UhE7ZjTilKBqauLAV9GJMGcVBN_o2OKMlNe605D1hkuPVUNJPAxRmy1ZEs6I7uv3uV5BqMJ2GSA",
    leave_balance:      18,
    carry_forward:      0,
  },
  {
    name:               "Tushar Thapliyal",
    email:              "tushar@8thelement.ai",
    role:               "manager",
    bot_role:           "hr",
    manager:            "",
    manager_email:      "",
    manager_teams_id:   "",
    teamlead:           "",
    teamlead_email:     "",
    teamlead_teams_id:  "",
    teams_id:           "29:1gCB8FOCd4_AdYy8lCvQBty0jW1Cs6xN1wnU4wMs2xJAQjesC2VhTR1PIt_Luo-yBLmHQ_PjAb0xg2ETM0mc-kg",
    leave_balance:      22,
    carry_forward:      0,
  },
  {
    name:               "devtools",
    email:              "devtools@company.com",
    role:               "employee",
    bot_role:           "hr",
    manager:            "Varsha M",
    manager_email:      "varsha.m@8thelement.ai",
    manager_teams_id:   "29:1PXmmVnyMqR2UhE7ZjTilKBqauLAV9GJMGcVBN_o2OKMlNe605D1hkuPVUNJPAxRmy1ZEs6I7uv3uV5BqMJ2GSA",
    teamlead:           "Varsha M",
    teamlead_email:     "varsha.m@8thelement.ai",
    teamlead_teams_id:  "29:1PXmmVnyMqR2UhE7ZjTilKBqauLAV9GJMGcVBN_o2OKMlNe605D1hkuPVUNJPAxRmy1ZEs6I7uv3uV5BqMJ2GSA",
    teams_id:           "devtools",
    leave_balance:      22,
    carry_forward:      0,
  },
];

// ── Holidays 2026 ──────────────────────────────────────────────────────────

const holidays2026 = [
  { date: "2026-01-01", name: "New Year's Day",     added_by: "seed" },
  { date: "2026-01-14", name: "Makara Sankranti",   added_by: "seed" },
  { date: "2026-01-26", name: "Republic Day",       added_by: "seed" },
  { date: "2026-03-03", name: "Holi",               added_by: "seed" },
  { date: "2026-03-19", name: "Ugadi",              added_by: "seed" },
  { date: "2026-03-20", name: "Idul Fitr / Ramzan", added_by: "seed" },
  { date: "2026-08-15", name: "Independence Day",   added_by: "seed" },
  { date: "2026-09-14", name: "Ganesh Chaturthi",   added_by: "seed" },
  { date: "2026-10-02", name: "Gandhi Jayanti",     added_by: "seed" },
  { date: "2026-10-20", name: "Vijaya Dashami",     added_by: "seed" },
  { date: "2026-11-08", name: "Diwali",             added_by: "seed" },
  { date: "2026-12-25", name: "Christmas Day",      added_by: "seed" },
];

// ── Seed ───────────────────────────────────────────────────────────────────

async function main() {
  console.log("═══════════════════════════════════");
  console.log("  LeaveAgent — Postgres Seed       ");
  console.log("═══════════════════════════════════\n");

  await prisma.$connect();
  console.log("✅ Connected to Neon Postgres\n");

  // Employees
  console.log("Seeding employees...");
  let empCount = 0;
  for (const emp of employees) {
    try {
      await prisma.employee.upsert({
        where:  { name: emp.name },
        update: emp,
        create: emp,
      });
      console.log(`  ✅ ${emp.name} (${emp.bot_role})`);
      empCount++;
    } catch (err: any) {
      console.error(`  ❌ ${emp.name}: ${err.message}`);
    }
  }
  console.log(`\nEmployees: ${empCount}/${employees.length} seeded\n`);

  // Holidays
  console.log("Seeding holidays...");
  let holCount = 0;
  for (const h of holidays2026) {
    try {
      await prisma.holiday.upsert({
        where:  { date: h.date },
        update: { name: h.name, added_by: h.added_by },
        create: h,
      });
      console.log(`  ✅ ${h.date} — ${h.name}`);
      holCount++;
    } catch (err: any) {
      console.error(`  ❌ ${h.date}: ${err.message}`);
    }
  }
  console.log(`\nHolidays: ${holCount}/${holidays2026.length} seeded\n`);

  console.log("═══════════════════════════════════");
  console.log("  Seed complete! ✅                ");
  console.log("═══════════════════════════════════");
  console.log("\nNext steps:");
  console.log("  1. npm run dev");
  console.log("  2. Each person messages bot once → conversation refs saved to Postgres");
  console.log("  3. Bot remembers everyone even after restarts ✅\n");
}

main()
  .catch((err) => { console.error("Seed failed:", err); process.exit(1); })
  .finally(() => prisma.$disconnect());
