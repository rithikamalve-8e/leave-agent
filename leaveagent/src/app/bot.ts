import {
  TeamsActivityHandler,
  TurnContext,
  MessageFactory,
  CardFactory,
} from "botbuilder";
import { parseLeaveIntent, ParsedLeaveIntent } from "./groqParser";

export class LeaveBot extends TeamsActivityHandler {
  constructor() {
    super();

    // Handle incoming messages from employees
    this.onMessage(async (context: TurnContext, next) => {
      const userMessage = context.activity.text?.trim() ?? "";
      const userName =
        context.activity.from.name ?? context.activity.from.id ?? "Employee";

      console.log(`[LeaveBot] Message from ${userName}: "${userMessage}"`);

      // Show typing indicator while processing
      await context.sendActivity({ type: "typing" });

      // Parse the intent via Groq AI
      const intent = await parseLeaveIntent(userMessage);

      console.log(`[LeaveBot] Parsed intent:`, JSON.stringify(intent, null, 2));

      // Route based on parsed result
      if (intent.needs_clarification || intent.intent === "UNKNOWN") {
        await handleClarification(context, intent);
      } else {
        await handleLeaveRequest(context, userName, intent);
      }

      await next();
    });

    // Handle members joining the bot chat
    this.onMembersAdded(async (context: TurnContext, next) => {
      const membersAdded = context.activity.membersAdded ?? [];
      for (const member of membersAdded) {
        if (member.id !== context.activity.recipient.id) {
          await context.sendActivity(
            MessageFactory.text(
              `👋 Hi! I'm **LeaveAgent**, your AI-powered leave assistant.\n\n` +
                `Just tell me what you need:\n` +
                `• \`WFH tomorrow\`\n` +
                `• \`Sick today\`\n` +
                `• \`Leave on Friday\`\n\n` +
                `I'll handle the rest! ✅`
            )
          );
        }
      }
      await next();
    });
  }
}

/**
 * Handles requests that need clarification from the employee.
 */
async function handleClarification(
  context: TurnContext,
  intent: ParsedLeaveIntent
): Promise<void> {
  const question =
    intent.clarification_question ??
    "Could you clarify your request? Try: 'WFH tomorrow' or 'Leave on Friday'.";

  await context.sendActivity(MessageFactory.text(`🤔 ${question}`));
}

/**
 * Handles a valid, fully-parsed leave or WFH request.
 */
async function handleLeaveRequest(
  context: TurnContext,
  userName: string,
  intent: ParsedLeaveIntent
): Promise<void> {
  const { intent: type, date, duration } = intent;

  // Format display date
  const displayDate = formatDisplayDate(date);
  const durationLabel = duration === "half_day" ? "Half Day" : "Full Day";
  const typeLabel = getTypeLabel(type);
  const typeEmoji = getTypeEmoji(type);

  // Build confirmation card
  const card = CardFactory.adaptiveCard({
    $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
    type: "AdaptiveCard",
    version: "1.4",
    body: [
      {
        type: "TextBlock",
        text: `${typeEmoji} Request Received`,
        weight: "Bolder",
        size: "Large",
        color: "Accent",
      },
      {
        type: "FactSet",
        facts: [
          { title: "Employee", value: userName },
          { title: "Type", value: typeLabel },
          { title: "Date", value: displayDate },
          { title: "Duration", value: durationLabel },
          { title: "Status", value: "⏳ Pending Manager Approval" },
        ],
      },
      {
        type: "TextBlock",
        text: "Your manager has been notified and will review shortly.",
        wrap: true,
        color: "Good",
        size: "Small",
      },
    ],
  });

  await context.sendActivity(MessageFactory.attachment(card));

  console.log(
    `[LeaveBot] ✅ Request confirmed for ${userName}: ${type} on ${date}`
  );

  // TODO (Day 2): Look up manager from Excel and send approval card
  // TODO (Day 3): Write to LeaveRequests.xlsx and post Teams announcement
}

/**
 * Format ISO date string to human-readable label.
 */
function formatDisplayDate(isoDate: string): string {
  if (!isoDate) return "Unknown date";
  try {
    const d = new Date(isoDate);
    return d.toLocaleDateString("en-IN", {
      weekday: "long",
      year: "numeric",
      month: "long",
      day: "numeric",
    });
  } catch {
    return isoDate;
  }
}

function getTypeLabel(intent: string): string {
  switch (intent) {
    case "WFH":
      return "Work From Home";
    case "LEAVE":
      return "Planned Leave";
    case "SICK":
      return "Sick Leave";
    default:
      return "Leave Request";
  }
}

function getTypeEmoji(intent: string): string {
  switch (intent) {
    case "WFH":
      return "🏠";
    case "LEAVE":
      return "🌴";
    case "SICK":
      return "🤒";
    default:
      return "📋";
  }
}