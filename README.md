# LeaveAgent — AI-Powered Leave & WFH Management Bot for Microsoft Teams

LeaveAgent is a Microsoft Teams bot that automates leave and WFH requests using natural language. Employees simply type messages like "WFH tomorrow" or "Sick today" and the bot handles parsing, approval routing, notifications, email alerts, and Excel logging — end to end.

---

## How It Works

```
Employee types: "Leave from 20th to 25th March"
        ↓
Bot parses intent using OpenAI
        ↓
Employee sees a Preview Card — Confirm / Edit / Cancel
        ↓
On Confirm → request saved to Excel
        ↓
Approval card sent to Team Lead (if employee) or Manager (if team lead)
        ↓
Email sent to approver + CC list
        ↓
Approver clicks Approve / Reject
        ↓
Employee gets notified in their DM
        ↓
Announcement sent to group channel
        ↓
Leave balance updated in Excel
```

---

## Project Structure

```
leaveagent/
├── src/
│   ├── index.ts                  ← Main bot entry point
│   └── app/
│       ├── groqParser.ts         ← AI intent parsing via Groq
│       ├── excelManager.ts       ← Excel read/write + in-memory state
│       ├── cards.ts              ← All Adaptive Card builders
│       └── emailService.ts       ← Outlook SMTP email notifications
├── scripts/
│   └── seedExcel.ts              ← Creates sample Employees.xlsx + LeaveRequests.xlsx
├── data/
│   ├── Employees.xlsx            ← Employee directory (edit this)
│   └── LeaveRequests.xlsx        ← Auto-created on first request
├── appPackage/
│   ├── manifest.json             ← Teams app manifest
│   ├── color.png                 ← 192x192 bot icon
│   └── outline.png               ← 32x32 bot icon
├── env/
│   └── .env.local                ← All environment variables (never commit this)
└── package.json
```

---

## Tech Stack

| Layer | Technology |
|-------|-----------|
| Bot Framework | `@microsoft/teams.apps` v2 (Teams SDK) |
| AI Parsing | `groq-sdk` — llama-3.1-8b-instant |
| Adaptive Cards | `@microsoft/teams.cards` v2 |
| Excel Storage | `xlsx` |
| Email | `nodemailer` via Office 365 SMTP |
| Local Dev Tunnel | ngrok |
| Runtime | Node.js 20/22 + TypeScript |

---

## Environment Variables

File: `env/.env.local`

```env
# Bot credentials (from Azure App Registration)
BOT_ID=your-azure-app-id
BOT_PASSWORD=your-azure-client-secret-value

# Groq AI
GROQ_API_KEY=your-groq-api-key

# Email (Outlook / Office 365)
EMAIL_USER=youremail@company.com
EMAIL_PASS=your-outlook-app-password

# HR email (global CC on all leave emails)
HR_EMAIL=hr@company.com

# Teams channel/group for announcements (get by adding bot to group)
ANNOUNCEMENT_CHANNEL_ID=19:abc123...

# Port (default 3978)
PORT=3978
```

> `BOT_PASSWORD` must be the **Value** of the client secret, not the Secret ID.
> Get an Outlook App Password from: Microsoft Account → Security → Advanced Security → App Passwords.

---

## Employees.xlsx Schema

Located at `data/Employees.xlsx`. Edit this file to add your team.

| Column | Description |
|--------|-------------|
| `name` | Full name exactly as it appears in Teams |
| `email` | Work email |
| `role` | `employee` or `teamlead` |
| `manager` | Manager's display name |
| `manager_email` | Manager's email |
| `manager_teams_id` | Manager's Teams ID (29:1abc...) |
| `teamlead` | Team lead's display name (blank if role = teamlead) |
| `teamlead_email` | Team lead's email |
| `teamlead_teams_id` | Team lead's Teams ID |
| `teams_id` | Employee's own Teams ID (29:1abc...) |
| `leave_balance` | Annual leave days remaining (default: 22) |

### How to get Teams IDs

**Option 1 — From terminal logs (recommended):**
Have each person send `help` to the bot. Terminal prints:
```
[LeaveAgent] "help" from Rithika MR (29:1zhcz...)
```
Copy the ID in brackets.

**Option 2 — Microsoft Graph Explorer:**
1. Go to https://developer.microsoft.com/graph/graph-explorer
2. Sign in with work account
3. Run: `GET https://graph.microsoft.com/v1.0/users/person@company.com`
4. Copy the `id` field from the response

---

## Approval Chain Logic

| Employee Role | Approval Card Sent To | Email To | Email CC |
|--------------|----------------------|----------|----------|
| `employee` | Team Lead | Team Lead | Manager + HR |
| `teamlead` | Manager | Manager | HR only |

On decision (Approved / Rejected):

| Employee Role | Notification To | CC |
|--------------|----------------|----|
| `employee` | Employee DM | Manager + HR |
| `teamlead` | Employee DM | HR only |

---

## Leave Types Supported

| Type | Triggered By |
|------|-------------|
| WFH | "wfh", "work from home", "remote" |
| Sick Leave | "sick", "unwell", "fever", "doctor" |
| Planned Leave | "leave", "day off", "vacation", "PTO" |
| Maternity Leave | "maternity", "expecting", "pregnancy leave" |
| Paternity Leave | "paternity", "father leave", "newborn" |
| Adoption Leave | "adoption leave" |
| Marriage Leave | "marriage leave", "wedding leave" |

---

## Leave Policy (built into AI parser)

| Type | Entitlement | Notes |
|------|------------|-------|
| Annual Leave | 22 days/year | 18 accrued (1.5/month) + 4 mandatory Dec 25–31 |
| Sick Leave | No fixed limit | No advance notice required |
| Maternity | 16 weeks | Apply ≥12 weeks before due date |
| Paternity | 5 days | Within 2 weeks before to 4 weeks after birth |
| Marriage | 5 days | Must request ≥42 days in advance |
| Adoption | Case by case | HR reviews |

> Dec 25–31 is mandatory company leave — bot blocks requests for these dates.
> Leave balance is deducted from Excel automatically on approval.
> WFH does not consume leave balance.

---

## Bot Commands

| Command | Description |
|---------|-------------|
| `WFH tomorrow` | Work from home request |
| `Sick today` | Sick leave |
| `Leave from 20th to 25th` | Multi-day planned leave |
| `Leave on Friday` | Single day leave |
| `my requests` | View last 5 requests |
| `my balance` / `leave balance` | Check remaining leave days |
| `summary` | Today's workforce availability |
| `help` | Show all commands |

---

## Preview / Edit / Cancel Flow

After the AI parses a request, the employee sees a **Preview Card** before it's sent to the approver:

```
[Review Your Request]
Type:         Work From Home
Date:         Monday, 16 March 2026
Working Days: 1 day(s)
Duration:     Full Day

[Confirm & Send]  [Edit]  [Cancel]
```

- **Confirm & Send** → saves to Excel, sends approval card to approver
- **Edit** → bot asks what to change, employee rephrases, new preview shown
- **Cancel** → request discarded, nothing saved

---

## Running Locally (Development)

### Prerequisites
- Node.js 20 or 22
- ngrok (free account at ngrok.com)
- BOT_ID and BOT_PASSWORD from Azure App Registration
- GROQ_API_KEY from console.groq.com (free)

### Steps

**1. Install dependencies:**
```bash
npm install
```

**2. Seed Excel files:**
```bash
npx ts-node scripts/seedExcel.ts
```

**3. Fill in `data/Employees.xlsx`** with real names, emails, Teams IDs, and leave balances.

**4. Fill in `env/.env.local`** with all credentials.

**5. Start ngrok** (keep running in a separate terminal):
```bash
ngrok http 3978
```
Or with static domain (URL never changes):
```bash
ngrok http --domain=your-static-domain.ngrok-free.app 3978
```

**6. Update messaging endpoint** in Bot Framework (dev.botframework.com) → your bot → Settings:
```
https://your-ngrok-url.ngrok-free.app/api/messages
```

**7. Start the bot:**
```bash
npm run dev
```

**8. Upload app to Teams:**
```powershell
cd appPackage
Compress-Archive -Path manifest.json, color.png, outline.png -DestinationPath ../leaveagent.zip
```
Teams → Apps → Manage your apps → Upload a custom app → select `leaveagent.zip`

---

## Azure App Registration Setup

1. Go to portal.azure.com → App Registrations → find `leaveagent`
2. Copy **Application (client) ID** → this is your `BOT_ID`
3. Go to **Certificates & secrets** → New client secret → copy the **Value** → this is your `BOT_PASSWORD`
4. The secret Value is shown only once — save it immediately

> If the secret expires or is lost, create a new one and update `BOT_PASSWORD` in `.env.local`.

---

## Bot Framework Registration

1. Go to dev.botframework.com → your bot → Settings
2. Set **Messaging endpoint**: `https://your-ngrok-url/api/messages`
3. Under **Channels** → ensure **Microsoft Teams** is listed as Running

---

## Getting the Announcement Channel ID

1. Add LeaveAgent to the Teams channel or group chat where announcements should appear
2. Anyone in that group sends a message to the bot
3. Terminal prints the `conversationId` — copy it
4. Set `ANNOUNCEMENT_CHANNEL_ID=19:abc123...` in `.env.local`
5. Restart bot

After this, every approved leave/WFH will post an announcement card to that channel automatically.

---

## Excel for HR Access

To give HR real-time access to `LeaveRequests.xlsx`:

1. Move `data/LeaveRequests.xlsx` to a OneDrive shared folder
2. Update the path in `src/app/excelManager.ts`:
```typescript
const LEAVE_PATH = "C:/Users/YourName/OneDrive/SharedFolder/LeaveRequests.xlsx";
```
3. Share the OneDrive folder with HR
4. HR opens the file from SharePoint/OneDrive — sees live data

No code changes needed beyond the path update.

---

## Known Behaviours

- **Bot must be running + ngrok active** for Teams messages to work. If bot is stopped, messages queue and deliver when restarted.
- **Pending state is in-memory** — if bot restarts during a preview (before employee confirms), the pending request is lost. Employee must re-submit.
- **Leave balance deducts on approval**, not on submission.
- **Conversation refs are in-memory** — if bot restarts, proactive DMs won't work until each person messages the bot again.
- **Action.Submit buttons** in real Teams sometimes route through the `message` handler instead of `card.action` — handled via intercept block in `index.ts`.
- **ngrok free tier** gives a random URL on every restart — update the messaging endpoint each time, or use a static domain.

---

## Files You Should NOT Modify

| File | Reason |
|------|--------|
| `src/app/instructions.txt` | Teams Toolkit generated, unused by LeaveAgent |
| `teamsapp.local.yml` | Teams Toolkit provisioning config, F5 only |

## Files Safe to Delete

| File | Reason |
|------|--------|
| `src/app/app.ts` | Teams Toolkit auto-generated starter template, not used |
| `.localConfigs` | Generated by Teams Toolkit for F5, not needed for manual run |

---

## Quick Reference — Common Issues

| Problem | Fix |
|---------|-----|
| Bot not responding in Teams | Check ngrok is running + `npm run dev` is running |
| 401 Authorization error | Wrong `BOT_PASSWORD` — use the Value not the Secret ID |
| "Error encountered while rendering" | Proactive send missing `from`/`conversation`/`recipient` fields |
| Approve button says "I can only help with leave requests" | Card action routed as message — handled by intercept block in `index.ts` |
| end_date wrong by 1 day | LLM treating range as exclusive — fixed in groqParser prompt |
| Bot installed in Teams but silent | Messaging endpoint not set or pointing to stale ngrok URL |
| Notification going to approver instead of employee | Employee and approver have same `conversationId` — check `employee?.teams_id !== activity.from.id` condition |
| Emails not sending | Check `EMAIL_PASS` is an App Password, not your login password |
| Leave balance not deducting | Only deducts on Approval, not submission. Check `updateLeaveStatus` is being called |