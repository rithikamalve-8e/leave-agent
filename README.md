# LeaveAgent — Complete Technical Documentation

---

## 1. HIGH LEVEL DESIGN (HLD)

```
┌─────────────────────────────────────────────────────────────────────────────┐
│                           MICROSOFT TEAMS                                   │
│                                                                             │
│   Employee DM    Approver DM    HR DM      Group Channel                   │
│   (a:1abc...)    (a:1def...)  (a:1ghi...)  (19:xyz...)                     │
└────────┬──────────────┬──────────┬───────────────┬──────────────────────────┘
         │              │          │               │
         └──────────────┴──────────┴───────────────┘
                              │ HTTPS
                              ▼
                    ┌─────────────────┐
                    │   ngrok Tunnel   │
                    │ (static domain)  │
                    └────────┬────────┘
                             │ POST /api/messages
                             ▼
┌─────────────────────────────────────────────────────────────────────────────┐
│                        LeaveAgent Bot (Node.js + TypeScript)                │
│                                                                             │
│  ┌─────────────┐  ┌──────────────┐  ┌──────────────┐  ┌─────────────────┐ │
│  │  index.ts   │  │commandRouter │  │  handlers/   │  │notificationSvc  │ │
│  │  (entry)    │→ │  .ts         │→ │  employee    │  │  (proactive DMs │ │
│  │             │  │              │  │  approver    │  │   + announce)   │ │
│  │             │  │              │  │  hr          │  │                 │ │
│  └──────┬──────┘  └──────────────┘  └──────────────┘  └─────────────────┘ │
│         │                                                                   │
│  ┌──────▼───────────────────────────────────────────────────────────────┐  │
│  │                         Core Services                                │  │
│  │                                                                      │  │
│  │  groqParser.ts         postgresManager.ts      cards.ts             │  │
│  │  (OpenAI GPT-4o        (Prisma + Neon DB       (Adaptive Card       │  │
│  │   intent parsing)       all data ops)           builders)           │  │
│  │                                                                      │  │
│  │  roleGuard.ts          schedulers.ts                                │  │
│  │  (role-based           (9am summary,                                │  │
│  │   permissions)          month-end,                                  │  │
│  │                         Dec 25 carry)                               │  │
│  └──────────────────────────────────────────────────────────────────────┘  │
└──────────────────────────┬──────────────────────────────────────────────────┘
                           │
           ┌───────────────┼───────────────┐
           ▼               ▼               ▼
  ┌────────────────┐ ┌──────────┐ ┌──────────────┐
  │  Neon Postgres │ │ OpenAI   │ │  Azure AD    │
  │  (Prisma ORM)  │ │  API     │ │  (Bot Auth)  │
  │                │ │ GPT-4o   │ │              │
  │  Employees     │ └──────────┘ └──────────────┘
  │  LeaveRequests │
  │  ConvRefs      │
  │  PendingReqs   │
  │  Holidays      │
  │  AuditLog      │
  │  MonthlySummry │
  └────────────────┘
```

---

## 2. LEAVE BALANCE SYSTEM — COMPLETE LOGIC

### 2.1 Annual Entitlement

```
Total Annual Leave = 22 days

Breakdown:
  18 days  →  Accrued monthly at 1.5 days/month (Jan–Dec)
   4 days  →  Mandatory company shutdown (Dec 25–31, pre-assigned)

Monthly Accrual Rate: 1.5 days per calendar month
```

### 2.2 Balance Calculation Formula

```
Current Balance = carry_forward + (currentMonth × 1.5) - totalUsed + totalLOP

Where:
  carry_forward  = days carried from previous year (max 6, calculated Dec 25)
  currentMonth   = current month number (1–12)
  totalUsed      = sum of approved LEAVE + SICK days taken this year (up to today)
  totalLOP       = sum of lop_days on those same approved records (LOP adds back)

Example (March, no carry forward, 2 days sick taken):
  Balance = 0 + (3 × 1.5) - 2 + 0
          = 0 + 4.5 - 2
          = 2.5 days
```

### 2.3 Balance Consuming Leave Types

```
CONSUMES balance:  LEAVE, SICK
DOES NOT consume:  WFH, MATERNITY, PATERNITY, MARRIAGE, ADOPTION
```

### 2.4 LOP (Loss of Pay) Calculation

LOP happens when the employee requests more days than their current balance.

```
Single-month request:
  balance  = getLeaveBalance(employee)    → e.g. 1.5 days
  requested = 3 days
  granted   = MIN(3, 1.5) = 1.5 days
  lop       = MAX(0, 3 - 1.5) = 1.5 days (Loss of Pay)
  hasLop    = true

LOP days are stored in lop_days column on the LeaveRequest record.
On approval: deductLeaveBalance(employee, days_count)
             The lop_days are stored but do NOT get deducted again —
             they are added BACK in getLeaveBalance formula (+totalLOP)
             because LOP days did not actually come from the leave balance.
```

### 2.5 Cross-Month LOP Calculation

When a leave request spans multiple calendar months (e.g. Jan 29 – Feb 5):

```
Step 1: Split working days by month
  Jan: 2 working days
  Feb: 5 working days

Step 2: Check balance month by month
  Current balance (at time of request): 1.5 days

  Jan (current month):
    balance = 1.5
    granted = MIN(2, 1.5) = 1.5
    lop     = MAX(0, 2 - 1.5) = 0.5
    runningBalance = MAX(0, 1.5 - 2) = 0

  Feb (future month — accrual hasn't happened yet):
    balance = 0   ← future months always get 0 balance
    granted = MIN(5, 0) = 0
    lop     = MAX(0, 5 - 0) = 5
    runningBalance = 0

  Result:
    totalDays    = 7
    totalGranted = 1.5
    totalLOP     = 5.5
    splits = [
      { month: "01/2026", days: 2, balance: 1.5, granted: 1.5, lop: 0.5 },
      { month: "02/2026", days: 5, balance: 0,   granted: 0,   lop: 5   },
    ]
```

### 2.6 January Leave (Special Case)

If requesting January leave BEFORE Dec 25 carry-forward is calculated:

```
Condition: leaveMonth === 1 AND leaveYear > currentYear AND today < Dec 25
Result: needsCarryForward = true
        balance = 0, granted = 0, lop = all days requested
Bot message: "Please request January leaves after December 25th"
```

---

## 3. CARRY FORWARD — YEARLY LOGIC

### 3.1 When it runs

```
Trigger: December 25 at 9:00 AM (scheduled by schedulers.ts)
Job:     runCarryForward() → called by scheduleDec25CarryForward()
```

### 3.2 Calculation

```
For each employee:
  dec31Balance = getLeaveBalance(employee.name, date: Dec 31 of current year)
  carry_forward = MIN(dec31Balance, 6)

  → Maximum 6 days can be carried forward
  → Any balance above 6 is forfeited
  → Stored in Employee.carry_forward column in Postgres
```

### 3.3 How carry_forward feeds into next year

```
Jan 1 onwards:
  getLeaveBalance(employee) uses:
    carry_forward (now populated from Dec 25 job)
    + (currentMonth × 1.5)
    - totalUsed
    + totalLOP

So on Jan 1:
  Balance = carry_forward + (1 × 1.5) - 0 + 0
          = carry_forward + 1.5

On Feb 1:
  Balance = carry_forward + (2 × 1.5) - totalUsedInJan + totalLOPInJan
```

### 3.4 Timeline

```
Jan 1         → New year starts, balance = carry_forward + 1.5
Every month   → Balance grows by 1.5 automatically (formula-based, no job needed)
Dec 25        → carry_forward calculated and stored for ALL employees
Dec 25–31     → Mandatory company shutdown (no leave requests accepted)
Jan 1 next yr → New cycle begins using new carry_forward
```

---

## 4. MONTHLY SUMMARY — HOW IT WORKS

### 4.1 What it stores

```
One row per employee per month in MonthlySummary table:

month      = "03/2026"
employee   = "Rithika MR"
opening    = balance at START of month (before this month's accrual)
available  = opening + 1.5  (after this month's 1.5 accrual)
leaves     = approved LEAVE+SICK days taken this month
wfh        = approved WFH days this month
lop        = LOP days this month
closing    = available - leaves + lop
pending    = pending days snapshot at month end
```

### 4.2 When it runs

```
Trigger: Last working day of month at 6:00 PM
         (same time as HR takeover job)
         Also triggered manually: "build summary MM YYYY"
```

### 4.3 Formula

```
opening   = carry + ((monthNum - 1) × 1.5) - totalUsedBefore + totalLOPBefore
available = opening + 1.5
closing   = available - leavesThisMonth + lopThisMonth
```

---

## 5. COMPLETE MESSAGE FLOW — ALL SCENARIOS

### 5.1 Employee Submits Leave Request

```
Employee types: "sick tomorrow"
                      │
                      ▼
              saveConversationRef()
              (save to Postgres ConversationRef table)
                      │
                      ▼
              activityValue intercept?  ──NO──▶  continue
                      │
                      ▼
              getPendingRequest(userId)
              edit mode? history.length > 0?  ──NO──▶  continue
                      │
                      ▼
              routeCommand() → no command matched
                      │
                      ▼
              handleLeaveRequest()
                      │
                      ├─▶ onBehalfMatch check → block if not HR
                      │
                      ├─▶ send("Processing your request...")
                      │
                      ├─▶ parseLeaveIntent("sick tomorrow")
                      │         │
                      │         ▼
                      │    OpenAI GPT-4o
                      │    System prompt with today's date, policy rules
                      │    Returns: { intent: "SICK", date: "2026-03-26",
                      │               duration: "full_day", confidence: 1.0 }
                      │         │
                      │         ▼
                      │    validateParsedIntent()
                      │    - confidence check (< 0.6 → clarify)
                      │    - month rollover correction
                      │    - past date check
                      │    - weekend check
                      │    - Dec 25-31 block
                      │    - multi_day end_date check
                      │
                      ├─▶ isDuplicateRequest() → already has request for this date?
                      │
                      ├─▶ isOverlappingLeave() → date within existing window?
                      │
                      ├─▶ isHoliday() → is this a public holiday?
                      │
                      ├─▶ findEmployee(userName) → get employee from Postgres
                      │
                      ├─▶ countWorkingDays() → exclude weekends + holidays
                      │
                      ├─▶ checkLeaveBalance()
                      │    → single month or cross-month LOP calculation
                      │    → needsCarryForward check for January
                      │
                      ├─▶ savePendingRequest() → stored in Postgres PendingRequest
                      │   (NOT yet in LeaveRequests table)
                      │
                      └─▶ send(buildPreviewCard())
                             [Confirm & Send] [Edit] [Cancel]
```

### 5.2 Employee Confirms Preview

```
Employee clicks "Confirm & Send"
                      │
            card.action: preview_confirm
                      │
                      ▼
              submitRequest()
                      │
                      ├─▶ getPendingRequest(userId) → get from Postgres
                      │
                      ├─▶ addLeaveRequest() → save to LeaveRequests table
                      │                       status: "Pending"
                      │
                      ├─▶ clearPendingRequest(userId)
                      │
                      ├─▶ send(buildConfirmationCard())
                      │   "Request Submitted — Awaiting approval"
                      │
                      ├─▶ sendHRAlert("submitted") → HR gets alert card
                      │
                      └─▶ sendApprovalCardToApprover()
                              │
                              ├─▶ getConversationRef(approverTeamsId)
                              │
                              ├─▶ api.conversations.activities().create()
                              │   → proactive DM to approver
                              │
                              └─▶ fallback: send inline if no ref
```

### 5.3 Approver Approves

```
Approver clicks "Approve" on card
                      │
            card.action handler (or message intercept)
                      │
                      ▼
              action === "reject"?
                      │
                   YES▼
              Return buildRejectionReasonPromptCard()
              (Input.Text card for rejection reason)
                      │
                   NO (approve)
                      ▼
              updateLeaveStatus(employee, date, "Approved", approverName)
                      │
                      ├─▶ Postgres: LeaveRequest.status = "Approved"
                      │            LeaveRequest.approved_by = approverName
                      │
                      ├─▶ deductLeaveBalance() if LEAVE or SICK type
                      │   Postgres: Employee.leave_balance -= days_count
                      │
                      ├─▶ sendStatusCardToEmployee()
                      │   "Your request was Approved ✅"
                      │   → proactive DM to employee's personal conv
                      │
                      ├─▶ sendApprovalAnnouncement()
                      │   "📅 Rithika MR will be on sick leave tomorrow."
                      │   → plain text to announcement group channel
                      │
                      ├─▶ sendWorkforceCardToManager()
                      │   Workforce Availability Update card
                      │   → proactive DM to manager above approver
                      │
                      ├─▶ sendHRAlert("approved")
                      │   → proactive DM to HR
                      │
                      └─▶ Return buildApprovedCardContent()
                          (replaces approval card in approver's chat)
```

### 5.4 Approver Rejects with Reason

```
Approver clicks "Reject"
                      │
                      ▼
              Return buildRejectionReasonPromptCard()
              [Input field for reason] [Confirm] [Cancel]
                      │
              Approver types reason and clicks "Confirm Rejection"
                      │
            card.action: confirm_reject
                      │
                      ▼
              updateLeaveStatus(employee, date, "Rejected", approverName, reason)
                      │
                      ├─▶ Postgres: status = "Rejected"
                      │            rejection_reason = reason
                      │            lop_days = 0 (reset)
                      │
                      ├─▶ sendStatusCardToEmployee()
                      │   "Your request was Rejected ❌"
                      │   Includes rejection reason
                      │
                      └─▶ Return buildRejectedCardContent(reason)
```

### 5.5 Edit Flow

```
Employee clicks "Edit" on preview card
                      │
            card.action: preview_edit
                      │
                      ▼
              getPendingRequest(userId)
              Set pending.history = [
                { role: "user",      content: "I want to request SICK on 2026-03-26" },
                { role: "assistant", content: "What would you like to change?" },
              ]
              savePendingRequest(userId, pending)
                      │
                      ▼
              send("What would you like to change?")
                      │
              Employee types: "change to leave on 28th"
                      │
            app.on("message")
                      │
                      ▼
              getPendingRequest(userId)
              history.length > 0 → edit mode
                      │
                      ▼
              handleEditMode()
                      │
                      ├─▶ parseLeaveIntent("change to leave on 28th", history)
                      │   → OpenAI sees full context
                      │   → Returns updated intent: LEAVE, date: 2026-03-28
                      │
                      ├─▶ recalculate daysCount, balanceResult
                      │
                      ├─▶ savePendingRequest() with history: [] (edit mode cleared)
                      │
                      └─▶ send(buildPreviewCard()) with new details
```

---

## 6. SCHEDULER FLOWS

### 6.1 Daily 9am Summary

```
Bot startup → scheduleDailySummary()
                      │
              msUntil(next 9am) calculated
              safeSetTimeout(fn, ms)
                      │
              At 9:00 AM (weekdays only):
                      │
                      ▼
              getTodaysAbsences()
              → Postgres: LeaveRequests WHERE date=today AND status=Approved
                      │
              Records found? → NO → skip (no message sent)
                      │
                   YES▼
              sendDailySummaryRest()
              "📋 Workforce Availability – 26 March: WFH: Rithika | Leave: Varsha"
              → Bot Framework REST API (OAuth token, no active request context)
              → Posted to ANNOUNCEMENT_CHANNEL_ID
                      │
              scheduleNext() → loop for next day
```

### 6.2 Month-End Approver Reminder (Last Working Day 9am)

```
Bot startup → scheduleApproverReminder(callback)
                      │
              Calculate lastWorkingDayOfMonth()
              Set target to 9:00 AM
              safeSetTimeout(fn, ms)
                      │
              On last working day at 9am:
                      │
                      ▼
              getMonthlyPendingRequests()
              → All Pending requests from current month
                      │
              Group by approver (using employee's teamlead/manager)
                      │
              For each approver with pending items:
                      │
                      ▼
              getConversationRef(approverTeamsId)
              sendCardViaRest() → buildApproverReminderCard()
              "⚠️ These requests from March are still pending — please action by EOD"
                      │
              scheduleNext() → next month
```

### 6.3 Month-End HR Takeover (Last Working Day 6pm)

```
Bot startup → scheduleHRTakeover(callback)
                      │
              Same last working day, but at 18:00
                      │
              On last working day at 6pm:
                      │
                      ▼
              getMonthlyPendingRequests()
              → Any STILL-pending requests (approver didn't act)
                      │
              Records found? → NO → skip
                      │
                   YES▼
              getConversationRef(HR_TEAMS_ID)
              sendCardViaRest() → buildHRTakeoverCard()
              "🚨 These requests were not actioned by approvers — please review"
                      │
              HR actions each request normally via bot commands
```

### 6.4 Dec 25 Carry Forward

```
Bot startup → scheduleDec25CarryForward()
              Target: Dec 25 at 9:00 AM
                      │
              On Dec 25 at 9am:
                      │
                      ▼
              runCarryForward()
                      │
              For each employee in Postgres:
                      │
                      ▼
              dec31Balance = getLeaveBalance(name, date: Dec 31)
                      │
              carry_forward = MIN(dec31Balance, 6)
              → Max 6 days, any excess is forfeited
                      │
              Postgres UPDATE employees SET carry_forward = value
                      │
              scheduleNext() → Dec 25 next year
```

---

## 7. ROLE SYSTEM — COMPLETE LOGIC

### 7.1 Role Detection

```
Every message → getRoleContext(userName)
                      │
                      ▼
              findEmployee(userName) in Postgres
                      │
              Read employee.bot_role field:
                "hr"       → BotRole.hr
                "approver" → BotRole.approver
                default    → BotRole.employee
```

### 7.2 Permission Matrix

```
Action                        Employee   Approver   HR
─────────────────────────────────────────────────────
Submit own leave                 ✅         ✅        ✅
View own requests                ✅         ✅        ✅
View own balance                 ✅         ✅        ✅
Delete own pending               ✅         ✅        ✅
Edit own pending                 ✅         ✅        ✅
Submit for others                ❌         ❌        ✅
Delete any request               ❌         ❌        ✅
Approve/Reject                   ❌         ✅*       ✅
View team requests               ❌         ✅        ✅
Query team leave                 ❌         ✅        ✅
View all org requests            ❌         ❌        ✅
Adjust leave balance             ❌         ❌        ✅
Add/edit/delete holidays         ❌         ❌        ✅
Download reports                 ❌         ❌        ✅
Approve own requests             ❌         ❌        ❌
Bypass policy checks             ❌         ❌        ✅

* Approver can only approve direct reportees, not their own requests
```

### 7.3 Approver Self-Submit Flow

```
Approver submits: "leave tomorrow"
                      │
                      ▼
              findEmployee(approverName)
              employee.role === "teamlead"?
                      │
                   YES▼
              approverTeamsId = employee.manager_teams_id
              approverName    = employee.manager
                      │
              → Approval card sent to THEIR manager, not themselves
```

---

## 8. DATA FLOW — POSTGRES TABLES

### 8.1 Table Relationships

```
Employee (1) ──────────────── (N) LeaveRequest
Employee (1) ──────────────── (1) ConversationRef
Employee (1) ──────────────── (N) MonthlySummary
Employee (1) ──────────────── (1) PendingRequest (ephemeral)

Holiday       (independent)
AuditLog      (independent)
```

### 8.2 Request Lifecycle in DB

```
1. Employee submits → PendingRequest created (preview stage)
   status: in-memory only (Postgres pending table)

2. Employee confirms → LeaveRequest created
   status: "Pending"
   PendingRequest deleted

3. Approver approves → LeaveRequest updated
   status: "Approved"
   approved_by: approverName
   Employee.leave_balance: -= days_count (if LEAVE/SICK)

4. Approver rejects → LeaveRequest updated
   status: "Rejected"
   rejection_reason: reason
   lop_days: 0 (reset)
   (leave_balance NOT deducted)

5. HR deletes → LeaveRequest soft-deleted
   status: "Deleted"
   deleted_by: hrName
   deleted_at: timestamp
   (AuditLog entry created)

6. HR restores → LeaveRequest restored
   status: "Pending"
   deleted_by: null
   deleted_at: null
```

---

## 9. WORKING DAYS CALCULATION

```
countWorkingDays(startDate, endDate):

  1. Load all holidays from Postgres Holiday table
  2. Build Set of holiday dates
  3. Iterate day by day from start to end:
     - Skip Saturday (day === 6)
     - Skip Sunday  (day === 0)
     - Skip if date in holiday set
     - Count remaining days
  4. Return { workingDays, holidays }

Example: March 19–25 2026
  March 19 = Thursday (Ugadi holiday) → skip
  March 20 = Friday   (Idul Fitr)      → skip
  March 21 = Saturday                  → skip
  March 22 = Sunday                    → skip
  March 23 = Monday                    → count (1)
  March 24 = Tuesday                   → count (2)
  March 25 = Wednesday                 → count (3)
  Result: 3 working days
```

---

## 10. OVERLAP / DUPLICATE DETECTION

### 10.1 Exact Duplicate

```
isDuplicateRequest(employee, date):
  SELECT * FROM leave_requests
  WHERE employee = ? AND date = ? AND status IN ('Pending', 'Approved')
  → If found → block
```

### 10.2 Window Overlap

```
isOverlappingLeave(employee, newStart, newEnd):
  SELECT all active requests (Pending + Approved) for employee
  For each existing request:
    existStart = record.date
    existEnd   = record.end_date ?? record.date

  Overlap if: newStart <= existEnd AND newEnd >= existStart

Example:
  Existing: 2026-04-15 to 2026-04-20
  New:      2026-04-17 to 2026-04-22
  → 17 <= 20 AND 22 >= 15 → OVERLAPS → blocked
```

---

## 11. CONVERSATION REF PERSISTENCE

```
Problem: Bot restarts wipe in-memory refs → proactive DMs fail

Solution: Store in Postgres ConversationRef table

On every message:
  saveConversationRef({
    userId, userName, conversationId,
    serviceUrl, tenantId, botId,
    isPersonal: conversationId.startsWith("a:")
  })

Priority logic:
  - If no existing ref → save
  - If new ref is personal (a:) → overwrite (preferred)
  - If existing is personal, new is group (19:) → keep existing
  
Result: Personal DM refs (a:...) always take priority over group refs (19:...)
        Bot restart → load from Postgres → proactive DMs work immediately
```

---

## 12. LOW LEVEL DESIGN (LLD)

### 12.1 File Structure

```
leaveagent/
├── src/
│   ├── index.ts                    ← Entry point, message + card.action handlers
│   └── app/
│       ├── groqParser.ts           ← OpenAI intent parsing + validation
│       ├── postgresManager.ts      ← All Prisma DB operations
│       ├── cards.ts                ← All Adaptive Card builders
│       ├── roleGuard.ts            ← Role detection + permission gates
│       ├── commandRouter.ts        ← Command table + routing
│       ├── notificationService.ts  ← All proactive DM + announcement sending
│       ├── schedulers.ts           ← Daily/monthly/yearly scheduled jobs
│       └── handlers/
│           ├── sharedHandlers.ts   ← help, summary, holidays (role-aware)
│           ├── employeeHandlers.ts ← my requests, balance, delete, edit
│           ├── approverHandlers.ts ← team requests, pending, who is on leave
│           └── hrHandlers.ts       ← all HR commands
├── prisma/
│   └── schema.prisma               ← DB schema (Neon Postgres)
├── scripts/
│   ├── seedPostgres.ts             ← Seed employees + holidays
│   └── migrateToPostgres.ts        ← One-time Excel → Postgres migration
└── env/
    └── .env.local                  ← All secrets (never commit)
```

### 12.2 groqParser.ts Internal Flow

```
parseLeaveIntent(userMessage, history?)
          │
          ▼
    buildSystemPrompt()
    → Inject today's date, current month, year
    → Include all policy rules, date rules, examples
          │
          ▼
    OpenAI GPT-4o call
    model: "gpt-4o"
    temperature: 0.1 (deterministic)
    max_tokens: 400
          │
          ▼
    Parse JSON response
          │
          ▼
    validateParsedIntent(raw, originalMessage)
          │
          ├─▶ confidence < 0.6 → needs_clarification
          ├─▶ month rollover correction
          │   (if LLM jumped to next month but user didn't say month name)
          ├─▶ auto end_date for PATERNITY/MARRIAGE/MATERNITY
          ├─▶ past date → block
          ├─▶ weekend → suggest Fri/Mon
          ├─▶ Dec 25-31 → block (mandatory shutdown)
          ├─▶ MARRIAGE < 42 days notice → warn
          ├─▶ multi_day missing end_date → ask
          └─▶ return ParsedLeaveIntent
```

### 12.3 commandRouter.ts Internal Flow

```
routeCommand(ctx, nctx, extras)
          │
          ▼
    Iterate commands[] array (ordered, first match wins):
          │
    For each CommandDef:
          │
          ├─▶ def.match(cmd, userMessage) → regex/string match
          │
          ├─▶ no match → next
          │
          └─▶ match found:
                    │
                    ├─▶ role check:
                    │   ctx.role.botRole === "hr" → always pass
                    │   else → check def.roles.includes(botRole)
                    │   not allowed → send permission error
                    │
                    └─▶ def.handler(ctx, nctx, extras)
                        return true (handled)
          │
    No match → return false → fall through to AI intent parsing
```

### 12.4 notificationService.ts — Two Modes

```
Mode 1: API mode (inside request handler)
  → Uses Teams SDK api object from active handler context
  → api.conversations.activities(convId).create({...})
  → Used for: approval card, status card, HR alert, workforce card

Mode 2: REST mode (inside schedulers)
  → No active request context, no api object
  → getBotToken() → Azure AD OAuth2 client_credentials
  → fetch(serviceUrl + /v3/conversations/{id}/activities)
  → Authorization: Bearer {token}
  → Used for: daily summary, approver reminders, HR takeover
```

### 12.5 Key Env Variables

```
BOT_ID / MICROSOFT_APP_ID        → Azure AD App Registration client ID
BOT_PASSWORD / MICROSOFT_APP_PASSWORD → Azure AD client secret value
MICROSOFT_APP_TENANT_ID          → Azure tenant ID
DATABASE_URL                     → Neon Postgres connection string
OPENAI_API_KEY                   → OpenAI API key (GPT-4o)
ANNOUNCEMENT_CHANNEL_ID          → Teams group chat conv ID (19:...)
HR_TEAMS_ID                      → HR person's Teams user ID (29:1...)
BOT_SERVICE_URL                  → Bot Framework service URL
PORT                             → Bot port (default 3978)
```

---

## 13. NOTIFICATION MATRIX (COMPLETE)

```
Event                    │ Employee      │ Approver      │ HR            │ Group
─────────────────────────┼───────────────┼───────────────┼───────────────┼──────────────
Request submitted        │ ✅ Confirm    │ ✅ Approval   │ ✅ Alert      │ ❌
                         │   card        │   card (DM)   │   card (DM)   │
─────────────────────────┼───────────────┼───────────────┼───────────────┼──────────────
Request approved         │ ✅ Status     │ ✅ Card       │ ✅ Alert      │ ✅ Text msg
                         │   card (DM)   │   updates     │   card (DM)   │   announcement
─────────────────────────┼───────────────┼───────────────┼───────────────┼──────────────
Request rejected         │ ✅ Status     │ ✅ Card       │ ❌            │ ❌
                         │   card +      │   updates     │               │
                         │   reason (DM) │               │               │
─────────────────────────┼───────────────┼───────────────┼───────────────┼──────────────
HR submits on behalf     │ ✅ Status     │ ❌            │ ❌            │ ✅ Text msg
                         │   card (DM)   │               │               │   announcement
─────────────────────────┼───────────────┼───────────────┼───────────────┼──────────────
HR adjusts balance       │ ✅ Balance    │ ❌            │ ❌            │ ❌
                         │   updated (DM)│               │               │
─────────────────────────┼───────────────┼───────────────┼───────────────┼──────────────
HR deletes request       │ ✅ Deleted    │ ✅ Deleted    │ ❌            │ ❌
                         │   notification│   notification│               │
─────────────────────────┼───────────────┼───────────────┼───────────────┼──────────────
Holiday added/changed    │ ✅ DM all     │ ✅ DM all     │ ❌            │ ✅ Text msg
─────────────────────────┼───────────────┼───────────────┼───────────────┼──────────────
Month-end (9am)          │ ❌            │ ✅ Reminder   │ ❌            │ ❌
                         │               │   card (DM)   │               │
─────────────────────────┼───────────────┼───────────────┼───────────────┼──────────────
Month-end (6pm)          │ ❌            │ ❌            │ ✅ Takeover   │ ❌
                         │               │               │   card (DM)   │
─────────────────────────┼───────────────┼───────────────┼───────────────┼──────────────
Daily 9am summary        │ ❌            │ ❌            │ ❌            │ ✅ Text msg
(weekdays only)          │               │               │               │   (if absent)
─────────────────────────┼───────────────┼───────────────┼───────────────┼──────────────
Approver on leave        │ ✅ Notified   │ ✅ Original + │ ❌            │ ❌
(escalation)             │   of escalate │   manager     │               │
```

---

## 14. LEAVE POLICY REFERENCE

```
Type            Entitlement        Balance?   Policy
────────────────────────────────────────────────────────────────────────────
LEAVE           22 days/year       YES        Dec 25-31 blocked (shutdown)
                (18 accrued +                  Max 6 days carry forward
                 4 mandatory Dec)              Up to 3 days extra carry (HR)

SICK            No fixed limit     YES        No advance notice required
                (deducted from                 Consumes annual balance
                 annual balance)

WFH             Unlimited          NO         No policy constraints

MATERNITY       16 weeks           NO         Apply ≥12 weeks before due date
                                              ≥80 days tenure required

PATERNITY       5 days             NO         Within 2 weeks before to
                                              4 weeks after birth
                                              Must provide start date

MARRIAGE        5 days             NO         ≥42 days advance notice required
                                              Post-probation only

ADOPTION        Case by case       NO         HR reviews individually
                                              Always needs_clarification: true

Dec 25-31       MANDATORY          N/A        Pre-assigned, no request needed
SHUTDOWN        company            (4 days)   Bot blocks any requests here
```
