import { type AnthropicProviderOptions, anthropic } from "@ai-sdk/anthropic";
import { devToolsMiddleware } from "@ai-sdk/devtools";
import type { InferAgentUIMessage } from "ai";
import { ToolLoopAgent, wrapLanguageModel } from "ai";
import * as z from "zod";
import { getSystemPrompt } from "@/server/ai/prompt";
import { tools } from "@/server/ai/tools";
import { Sheet } from "@/spreadsheet-service";

const wrappedAnthropic = (model: string) =>
  wrapLanguageModel({
    model: anthropic(model),
    middleware: devToolsMiddleware(),
  });

export const SpreadsheetAgent = new ToolLoopAgent({
  callOptionsSchema: z.object({
    model: z
      .enum(["claude-sonnet-4-5", "claude-opus-4-5"])
      .default("claude-opus-4-5"),
    sheets: z.array(Sheet),
  }),
  model: wrappedAnthropic("claude-opus-4-5"),
  tools,
  prepareCall: ({ options, ...initialOptions }) => {
    return {
      ...initialOptions,
      system: getSystemPrompt(options.sheets, "m"),
      model: wrappedAnthropic(options.model),
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
