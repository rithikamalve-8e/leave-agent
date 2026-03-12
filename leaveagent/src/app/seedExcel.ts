import * as XLSX from "xlsx";
import * as path from "path";
import * as fs   from "fs";

const DATA_DIR = path.join(process.cwd(), "data");
if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR);

const employees = [
  {
    name:               "devtools",
    email:              "devtools@company.com",
    role:               "employee",           // "employee" or "teamlead"
    manager:            "devtools",
    manager_email:      "devtools@company.com",
    manager_teams_id:   "devtools",
    teamlead:           "devtools",
    teamlead_email:     "devtools@company.com",
    teamlead_teams_id:  "devtools",           // for devtools, all same
    teams_id:           "devtools",
  },
  {
    name:               "Rahul",
    email:              "rahul@company.com",
    role:               "employee",           // Team Lead approves Rahul
    manager:            "Priya",
    manager_email:      "priya@company.com",
    manager_teams_id:   "REPLACE_PRIYA_TEAMS_ID",
    teamlead:           "Suresh",
    teamlead_email:     "suresh@company.com",
    teamlead_teams_id:  "REPLACE_SURESH_TEAMS_ID",
    teams_id:           "REPLACE_RAHUL_TEAMS_ID",
  },
  {
    name:               "Suresh",
    email:              "suresh@company.com",
    role:               "teamlead",           // Manager approves Suresh
    manager:            "Priya",
    manager_email:      "priya@company.com",
    manager_teams_id:   "REPLACE_PRIYA_TEAMS_ID",
    teamlead:           "",                   // not used for teamlead role
    teamlead_email:     "",
    teamlead_teams_id:  "",
    teams_id:           "REPLACE_SURESH_TEAMS_ID",
  },
];

const wb = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(
  wb,
  XLSX.utils.json_to_sheet(employees, {
    header: ["name","email","role","manager","manager_email","manager_teams_id","teamlead","teamlead_email","teamlead_teams_id","teams_id"],
  }),
  "Employees"
);
XLSX.writeFile(wb, path.join(DATA_DIR, "Employees.xlsx"));
console.log("Employees.xlsx created.");

const wb2 = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(
  wb2,
  XLSX.utils.json_to_sheet([], {
    header: ["employee","email","type","date","end_date","duration","status","approved_by","requested_at","updated_at"],
  }),
  "LeaveRequests"
);
XLSX.writeFile(wb2, path.join(DATA_DIR, "LeaveRequests.xlsx"));
console.log("LeaveRequests.xlsx created.");