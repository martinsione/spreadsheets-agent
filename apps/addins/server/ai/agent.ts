import {
  type AnthropicProviderOptions,
  createAnthropic,
} from "@ai-sdk/anthropic";
import { devToolsMiddleware } from "@ai-sdk/devtools";
import type { InferAgentUIMessage } from "ai";
import { ToolLoopAgent, wrapLanguageModel } from "ai";
import type * as z from "zod";
import { getSystemPrompt } from "@/server/ai/prompt";
import {
  callOptionsSchema,
  type messageMetadataSchema,
} from "@/server/ai/schema";
import { tools, writeTools } from "@/server/ai/tools";

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
  callOptionsSchema,
  prepareCall: ({ options, ...initialOptions }) => {
    const anthropic = createAnthropic({ apiKey: options.anthropicApiKey });
    const wrappedModel = wrapLanguageModel({
      model: anthropic(options.model.replace("anthropic:", "")),
      middleware: [
        //
        // devToolsMiddleware(), // <- needs to be removed for build
      ],
    });

    return {
      ...initialOptions,
      model: wrappedModel,
      system: getSystemPrompt(options.sheets, options.environment),
      providerOptions: {
        anthropic: {
          cacheControl: { type: "ephemeral" },
          thinking: { type: "enabled", budgetTokens: 16000 },
        } satisfies AnthropicProviderOptions,
      },
      headers: {
        "anthropic-beta":
          "interleaved-thinking-2025-05-14,fine-grained-tool-streaming-2025-05-14",
        "anthropic-dangerous-direct-browser-access": "true", // Needed to run directly in the browser
      },
    };
  },
});

export type SpreadsheetAgentUIMessage = InferAgentUIMessage<
  typeof SpreadsheetAgent,
  z.infer<typeof messageMetadataSchema>
>;
