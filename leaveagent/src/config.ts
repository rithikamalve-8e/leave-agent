import * as dotenv from "dotenv";
dotenv.config();

export const config = {
  // Microsoft Bot Framework credentials
  botId: process.env.BOT_ID ?? "",
  botPassword: process.env.BOT_PASSWORD ?? "",

  // Groq AI API key for intent parsing
  groqApiKey: process.env.GROQ_API_KEY ?? "",

  // Teams channel ID for workforce announcements (set in Day 3)
  announcementChannelId: process.env.ANNOUNCEMENT_CHANNEL_ID ?? "",

  // Path to Excel data files
  employeesFilePath: process.env.EMPLOYEES_FILE_PATH ?? "./data/Employees.xlsx",
  leaveRequestsFilePath:
    process.env.LEAVE_REQUESTS_FILE_PATH ?? "./data/LeaveRequests.xlsx",

  // App port
  port: parseInt(process.env.PORT ?? "3978", 10),
};

// Validate critical config on startup
export function validateConfig(): void {
  const missing: string[] = [];

  if (!config.botId) missing.push("BOT_ID");
  if (!config.botPassword) missing.push("BOT_PASSWORD");
  if (!config.groqApiKey) missing.push("GROQ_API_KEY");

  if (missing.length > 0) {
    console.warn(
      `[Config] ⚠️  Missing environment variables: ${missing.join(", ")}`
    );
    console.warn(`[Config] Add them to your .env.local file.`);
  } else {
    console.log("[Config] ✅ All required environment variables loaded.");
  }
}