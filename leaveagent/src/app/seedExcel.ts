import * as XLSX from "xlsx";
import * as path from "path";
import * as fs   from "fs";
 
const DATA_DIR = path.join(process.cwd(), "data");
if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR);
 
const employees = [
  {
    name:               "devtools",
    email:              "rithika.mr@8thelement.ai",
    role:               "employee",           // "employee" or "teamlead"
    manager:            "varsha",
    manager_email:      "varsha.m@8thelement.ai",
    manager_teams_id:   "28e6afa3-0515-4a5f-add5-9e70e3f68123",
    teamlead:           "varsha",
    teamlead_email:     "varsha.m@8thelement.ai",
    teamlead_teams_id:  "28e6afa3-56a5-4342-8d1b-4714fa77d546",           // for devtools, all same
    teams_id:           "a37249d6-7813-4258-922f-a7bc07291378",
    leave_balance:      22,  // annual leave days
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
    leave_balance:      22,
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
    leave_balance:      22,
  },
];
 
const wb = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(
  wb,
  XLSX.utils.json_to_sheet(employees, {
    header: ["name","email","role","manager","manager_email","manager_teams_id","teamlead","teamlead_email","teamlead_teams_id","teams_id","leave_balance"],
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