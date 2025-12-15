import { createAgentUIStreamResponse } from "ai";
import {
  SpreadsheetAgent,
  type SpreadsheetAgentUIMessage,
} from "@/server/ai/agent";
import type { Sheet } from "@/spreadsheet-service";

async function POST(req: Request) {
  const body = (await req.json()) as {
    messages: SpreadsheetAgentUIMessage[];
    model: string;
    ANTHROPIC_API_KEY: string;
    sheets: Sheet[];
  };

  if (!body.ANTHROPIC_API_KEY) {
    return new Response("API key is required", { status: 400 });
  }

  return createAgentUIStreamResponse({
    agent: SpreadsheetAgent,
    options: {
      model: body.model,
      ANTHROPIC_API_KEY: body.ANTHROPIC_API_KEY,
      sheets: body.sheets,
    },
    messages: body.messages,
    sendReasoning: true,
    sendSources: true,
  });
}

export function chatRoute(req: Request) {
  switch (req.method) {
    case "POST":
      return POST(req);
    default:
      return new Response("Method not allowed", { status: 405 });
  }
}
