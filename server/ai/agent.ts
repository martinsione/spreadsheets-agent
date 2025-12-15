import {
  type AnthropicProviderOptions,
  createAnthropic,
} from "@ai-sdk/anthropic";
import { devToolsMiddleware } from "@ai-sdk/devtools";
import type { InferAgentUIMessage } from "ai";
import { ToolLoopAgent, wrapLanguageModel } from "ai";
import * as z from "zod";
import { getSystemPrompt } from "@/server/ai/prompt";
import { tools, writeTools } from "@/server/ai/tools";
import { Sheet } from "@/spreadsheet-service";

const Models = z.enum(["claude-sonnet-4-5", "claude-opus-4-5"]);

const toolsWithApprovalRequiredConfigured = Object.fromEntries(
  Object.entries(tools).map(([name, toolDef]) => [
    name,
    writeTools.includes(name as (typeof writeTools)[number])
      ? { ...toolDef, needsApproval: true }
      : toolDef,
  ]),
) as typeof tools;

export const SpreadsheetAgent = new ToolLoopAgent({
  model: "", // Will be set in `prepareCall`
  tools: toolsWithApprovalRequiredConfigured,
  callOptionsSchema: z.object({
    model: Models.default("claude-opus-4-5"),
    ANTHROPIC_API_KEY: z.string(),
    sheets: z.array(Sheet),
  }),
  prepareCall: ({ options, ...initialOptions }) => {
    const anthropic = createAnthropic({ apiKey: options.ANTHROPIC_API_KEY });
    const wrappedModel = wrapLanguageModel({
      model: anthropic(options.model),
      middleware: devToolsMiddleware(),
    });

    return {
      ...initialOptions,
      model: wrappedModel,
      system: getSystemPrompt(options.sheets, "excel"),
      providerOptions: {
        anthropic: {
          thinking: { type: "enabled", budgetTokens: 16000 },
        } satisfies AnthropicProviderOptions,
      },
    };
  },
});

export type SpreadsheetAgentUIMessage = InferAgentUIMessage<
  typeof SpreadsheetAgent
>;
