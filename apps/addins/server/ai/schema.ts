import * as z from "zod";
import { Sheet } from "@/spreadsheet-service";

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

export const callOptionsSchema = z.object({
  anthropicApiKey: z.string(),
  model: modelSchema.default("anthropic:claude-opus-4-5"),
  sheets: z.array(Sheet),
  environment: z.enum(["excel", "google-sheets"]),
});

export const messageMetadataSchema = z.object({
  model: modelSchema.optional(),
  cachedInputTokens: z.number().optional(),
  inputTokens: z.number().optional(),
  outputTokens: z.number().optional(),
  reasoningTokens: z.number().optional(),
  totalTokens: z.number().optional(),
});
