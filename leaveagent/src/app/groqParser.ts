import Groq from "groq-sdk";

export interface ParsedLeaveIntent {
  intent: "WFH" | "LEAVE" | "SICK" | "UNKNOWN";
  date: string; // ISO format: YYYY-MM-DD
  end_date?: string; // For multi-day leave
  duration: "full_day" | "half_day" | "multi_day";
  reason?: string;
  needs_clarification: boolean;
  clarification_question?: string;
}

const client = new Groq({
  apiKey: process.env.GROQ_API_KEY,
});

const today = new Date().toISOString().split("T")[0];

const SYSTEM_PROMPT = `You are a leave request parser for a corporate HR bot.

Today's date is ${today}.

Your job: Convert an employee's natural language message into a structured JSON object.

Rules:
- intent must be one of: WFH, LEAVE, SICK, UNKNOWN
  - WFH = work from home
  - LEAVE = planned leave / day off
  - SICK = sick leave / unwell
  - UNKNOWN = cannot determine
- date must be a real calendar date in YYYY-MM-DD format
  - "tomorrow" = next calendar day
  - "today" = current date
  - "Friday" = the upcoming Friday
  - "next week" = Monday of next week
- duration: "full_day" unless they say "half day", "morning", "afternoon"
- needs_clarification: true only if the message is too vague to act on
- If needs_clarification is true, provide a clarification_question

Respond ONLY with valid JSON. No explanation, no markdown, no extra text.

Example input: "WFH tomorrow"
Example output:
{
  "intent": "WFH",
  "date": "2026-03-10",
  "duration": "full_day",
  "needs_clarification": false
}

Example input: "I need some time off"
Example output:
{
  "intent": "LEAVE",
  "date": "",
  "duration": "full_day",
  "needs_clarification": true,
  "clarification_question": "Sure! Which date would you like to take leave on?"
}`;

export async function parseLeaveIntent(
  userMessage: string
): Promise<ParsedLeaveIntent> {
  try {
    const completion = await client.chat.completions.create({
      model: "llama3-8b-8192",
      messages: [
        { role: "system", content: SYSTEM_PROMPT },
        { role: "user", content: userMessage },
      ],
      temperature: 0.1,
      max_tokens: 300,
    });

    const raw = completion.choices[0]?.message?.content?.trim() ?? "";

    // Strip markdown code fences if present
    const cleaned = raw.replace(/```json|```/g, "").trim();

    const parsed: ParsedLeaveIntent = JSON.parse(cleaned);
    return parsed;
  } catch (err) {
    console.error("[groqParser] Failed to parse intent:", err);
    return {
      intent: "UNKNOWN",
      date: "",
      duration: "full_day",
      needs_clarification: true,
      clarification_question:
        "Sorry, I didn't understand that. Could you say something like 'WFH tomorrow' or 'Leave on Friday'?",
    };
  }
}