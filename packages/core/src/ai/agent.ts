import {
  type AnthropicProviderOptions,
  createAnthropic,
} from "@ai-sdk/anthropic";
import type { InferAgentUIMessage, Tool } from "ai";
import { ToolLoopAgent, wrapLanguageModel } from "ai";
import * as z from "zod";
import { type MCPConnection, mergeMCPTools } from "./mcp";
import { getSystemPrompt } from "./prompt";
import { callOptionsSchema, type messageMetadataSchema } from "./schema";
import { tools, writeTools } from "./tools";

const toolsWithApprovalRequiredConfigured = Object.fromEntries(
  Object.entries(tools).map(([name, toolDef]) => [
    name,
    writeTools.includes(name as (typeof writeTools)[number])
      ? { ...toolDef, needsApproval: true }
      : toolDef,
  ]),
) as typeof tools;

/**
 * Extended call options schema that includes MCP connections.
 * MCP connections should be created before calling the agent and passed here.
 */
export const callOptionsWithMCPSchema = callOptionsSchema.extend({
  /** Pre-created MCP connections to use for tools */
  mcpConnections: z.array(z.custom<MCPConnection>()).optional().default([]),
});

/**
 * Type for the combined tools (base + MCP tools).
 * Uses intersection type to indicate additional dynamic tools may be present.
 */
type CombinedTools = typeof tools & Record<string, Tool>;

export const SpreadsheetAgent = new ToolLoopAgent({
  model: "", // Will be set in `prepareCall`
  tools: toolsWithApprovalRequiredConfigured,
  callOptionsSchema: callOptionsWithMCPSchema,
  prepareCall: ({ options, ...initialOptions }) => {
    const anthropic = createAnthropic({ apiKey: options.anthropicApiKey });
    const wrappedModel = wrapLanguageModel({
      model: anthropic(options.model.replace("anthropic:", "")),
      middleware: [],
    });

    // Merge MCP tools with existing tools if MCP connections are provided
    let finalTools = toolsWithApprovalRequiredConfigured as CombinedTools;
    if (options.mcpConnections && options.mcpConnections.length > 0) {
      finalTools = mergeMCPTools(
        toolsWithApprovalRequiredConfigured,
        options.mcpConnections,
        {
          prefixWithServerId: options.mcp?.prefixToolsWithServerId ?? false,
        },
      ) as CombinedTools;
    }

    return {
      ...initialOptions,
      model: wrappedModel,
      tools: finalTools as typeof tools,
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
