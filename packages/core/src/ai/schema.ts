import * as z from "zod";
import { Sheet } from "../spreadsheet-service";

export const models = [
  {
    name: "Claude Opus 4.5",
    value: "anthropic:claude-opus-4-5",
  },
  {
    name: "Claude Sonnet 4.5",
    value: "anthropic:claude-sonnet-4-5",
  },
] as const;

const modelSchema = z.enum(models.map((m) => m.value));

/**
 * MCP transport configuration schema.
 */
export const mcpTransportConfigSchema = z.object({
  type: z.enum(["sse", "http"]),
  url: z.string().url(),
  headers: z.record(z.string(), z.string()).optional(),
});

/**
 * MCP server configuration schema.
 */
export const mcpServerConfigSchema = z.object({
  id: z.string(),
  name: z.string(),
  transport: mcpTransportConfigSchema,
  toolsRequiringApproval: z.array(z.string()).optional(),
});

/**
 * MCP configuration schema.
 */
export const mcpConfigSchema = z.object({
  servers: z.array(mcpServerConfigSchema).default([]),
  /** If true, prefix MCP tool names with server ID to avoid conflicts */
  prefixToolsWithServerId: z.boolean().default(false),
});

export const callOptionsSchema = z.object({
  anthropicApiKey: z.string(),
  model: modelSchema.default("anthropic:claude-opus-4-5"),
  sheets: z.array(Sheet),
  environment: z.enum(["excel", "google-sheets", "web"]),
  /** MCP configuration for external tool servers */
  mcp: mcpConfigSchema.optional(),
});

export const messageMetadataSchema = z.object({
  model: modelSchema.optional(),
  cachedInputTokens: z.number().optional(),
  inputTokens: z.number().optional(),
  outputTokens: z.number().optional(),
  reasoningTokens: z.number().optional(),
  totalTokens: z.number().optional(),
});
