/**
 * scripts/migrateToPostgres.ts
 *
 * One-time migration script — reads Employees.xlsx and LeaveRequests.xlsx
 * and imports all data into Neon Postgres via Prisma.
 *
 * Run: npx ts-node scripts/migrateToPostgres.ts
 */

import * as dotenv from "dotenv";
import * as path   from "path";
dotenv.config({ path: path.resolve(process.cwd(), ".env") });

import * as XLSX   from "xlsx";
import * as fs     from "fs";
import { PrismaClient } from "@prisma/client";

const prisma = new PrismaClient();

const DATA_DIR       = path.join(process.cwd(), "data");
const EMPLOYEES_PATH = path.join(DATA_DIR, "Employees.xlsx");
const LEAVE_PATH     = path.join(DATA_DIR, "LeaveRequests.xlsx");
const CONV_REFS_PATH = path.join(DATA_DIR, "conversationRefs.json");

// ── Helpers ────────────────────────────────────────────────────────────────

function readSheet<T>(filePath: string, sheetIndex = 0): T[] {
  if (!fs.existsSync(filePath)) {
    console.warn(`[Migrate] File not found: ${filePath}`);
    return [];
  }
  const wb = XLSX.readFile(filePath);
  const ws = wb.Sheets[wb.SheetNames[sheetIndex]];
  return XLSX.utils.sheet_to_json<T>(ws);
}

// ── Migrate Employees ──────────────────────────────────────────────────────

async function migrateEmployees() {
  console.log("\n[Migrate] Importing employees...");

  const rows = readSheet<any>(EMPLOYEES_PATH);
  let imported = 0;
  let skipped  = 0;

  for (const row of rows) {
    const name = row.name?.toString().trim();
    if (!name || name.toLowerCase() === "devtools") {
      skipped++;
      continue;
    }

    try {
      await prisma.employee.upsert({
        where:  { name },
        update: {
          email:              row.email?.toString().trim()              ?? "",
          role:               row.role?.toString().trim()               ?? "employee",
          bot_role:           row.bot_role?.toString().trim()           ?? "employee",
          manager:            row.manager?.toString().trim()            ?? null,
          manager_email:      row.manager_email?.toString().trim()      ?? null,
          manager_teams_id:   row.manager_teams_id?.toString().trim()   ?? null,
          teamlead:           row.teamlead?.toString().trim()           ?? null,
          teamlead_email:     row.teamlead_email?.toString().trim()     ?? null,
          teamlead_teams_id:  row.teamlead_teams_id?.toString().trim()  ?? null,
          teams_id:           row.teams_id?.toString().trim()           || null,
          leave_balance:      Number(row.leave_balance)                 ?? 22,
        },
        create: {
          name,
          email:              row.email?.toString().trim()              ?? "",
          role:               row.role?.toString().trim()               ?? "employee",
          bot_role:           row.bot_role?.toString().trim()           ?? "employee",
          manager:            row.manager?.toString().trim()            ?? null,
          manager_email:      row.manager_email?.toString().trim()      ?? null,
          manager_teams_id:   row.manager_teams_id?.toString().trim()   ?? null,
          teamlead:           row.teamlead?.toString().trim()           ?? null,
          teamlead_email:     row.teamlead_email?.toString().trim()     ?? null,
          teamlead_teams_id:  row.teamlead_teams_id?.toString().trim()  ?? null,
          teams_id:           row.teams_id?.toString().trim()           || null,
          leave_balance:      Number(row.leave_balance)                 ?? 22,
        },
      });
      console.log(`  ✅ ${name}`);
      imported++;
    } catch (err: any) {
      console.error(`  ❌ ${name}: ${err.message}`);
      skipped++;
    }
  }

  console.log(`[Migrate] Employees: ${imported} imported, ${skipped} skipped`);
}

// ── Migrate Leave Requests ─────────────────────────────────────────────────

async function migrateLeaveRequests() {
  console.log("\n[Migrate] Importing leave requests...");

  const rows = readSheet<any>(LEAVE_PATH);
  let imported = 0;
  let skipped  = 0;

  for (const row of rows) {
    const employee = row.employee?.toString().trim();
    const date     = row.date?.toString().trim();

    if (!employee || !date) {
      skipped++;
      continue;
    }

    // Check employee exists in DB
    const emp = await prisma.employee.findFirst({
      where: { name: { equals: employee, mode: "insensitive" } },
    });

    if (!emp) {
      console.warn(`  ⚠️  Skipping request for unknown employee: ${employee}`);
      skipped++;
      continue;
    }

    // Check for duplicate
    const existing = await prisma.leaveRequest.findFirst({
      where: {
        employee: { equals: employee, mode: "insensitive" },
        date,
        type: row.type?.toString().trim() ?? "LEAVE",
      },
    });

    if (existing) {
      skipped++;
      continue;
    }

    try {
      await prisma.leaveRequest.create({
        data: {
          employee:         employee,
          email:            row.email?.toString().trim()        ?? emp.email,
          type:             row.type?.toString().trim()         ?? "LEAVE",
          date,
          end_date:         row.end_date?.toString().trim()     || null,
          duration:         row.duration?.toString().trim()     ?? "full_day",
          days_count:       Number(row.days_count)              ?? 1,
          reason:           row.reason?.toString().trim()       || null,
          rejection_reason: row.rejection_reason?.toString().trim() || null,
          status:           row.status?.toString().trim()       ?? "Pending",
          approved_by:      row.approved_by?.toString().trim()  || null,
          requested_at:     row.requested_at
                              ? new Date(row.requested_at)
                              : new Date(),
        },
      });
      console.log(`  ✅ ${employee} — ${row.type} on ${date}`);
      imported++;
    } catch (err: any) {
      console.error(`  ❌ ${employee} on ${date}: ${err.message}`);
      skipped++;
    }
  }

  console.log(`[Migrate] Leave requests: ${imported} imported, ${skipped} skipped`);
}

// ── Migrate Conversation Refs ──────────────────────────────────────────────

async function migrateConversationRefs() {
  console.log("\n[Migrate] Importing conversation refs...");

  if (!fs.existsSync(CONV_REFS_PATH)) {
    console.log("[Migrate] No conversationRefs.json found — skipping");
    return;
  }

  const raw  = fs.readFileSync(CONV_REFS_PATH, "utf-8");
  const refs = JSON.parse(raw) as Record<string, any>;
  let imported = 0;

  for (const [userId, ref] of Object.entries(refs)) {
    try {
      await prisma.conversationRef.upsert({
        where:  { userId },
        update: {
          userName:       ref.userName,
          conversationId: ref.conversationId,
          serviceUrl:     ref.serviceUrl,
          tenantId:       ref.tenantId ?? null,
          botId:          ref.botId,
          isPersonal:     ref.isPersonal ?? false,
        },
        create: {
          userId,
          userName:       ref.userName,
          conversationId: ref.conversationId,
          serviceUrl:     ref.serviceUrl,
          tenantId:       ref.tenantId ?? null,
          botId:          ref.botId,
          isPersonal:     ref.isPersonal ?? false,
        },
      });
      console.log(`  ✅ ${ref.userName}`);
      imported++;
    } catch (err: any) {
      console.error(`  ❌ ${ref.userName}: ${err.message}`);
    }
  }

  console.log(`[Migrate] Conversation refs: ${imported} imported`);
}

// ── Main ───────────────────────────────────────────────────────────────────

async function main() {
  console.log("═══════════════════════════════════════");
  console.log("  LeaveAgent → Neon Postgres Migration ");
  console.log("═══════════════════════════════════════");

  try {
    await prisma.$connect();
    console.log("[Migrate] Connected to Neon Postgres ✅");

    await migrateEmployees();
    await migrateLeaveRequests();
    await migrateConversationRefs();

    console.log("\n[Migrate] ✅ Migration complete!");
    console.log("[Migrate] Open Prisma Studio to verify: npx prisma studio");
  } catch (err) {
    console.error("[Migrate] Fatal error:", err);
    process.exit(1);
  } finally {
    await prisma.$disconnect();
  }
}

main();