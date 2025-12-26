import {
  createAgentUIStream as createAgentUIStreamBase,
  createUIMessageStreamResponse,
  smoothStream,
} from "ai";
import type * as z from "zod";
import {
  type callOptionsWithMCPSchema,
  SpreadsheetAgent,
  type SpreadsheetAgentUIMessage,
} from "./agent";
import {
  closeMCPConnections,
  createMCPConnections,
  type MCPConnection,
  type MCPServerConfig,
} from "./mcp";
import type { callOptionsSchema, messageMetadataSchema } from "./schema";
import type { tools } from "./tools";

export async function createAgentUIStream({
  body,
}: {
  body: {
    messages: SpreadsheetAgentUIMessage[];
    options: z.infer<typeof callOptionsSchema>;
  };
}) {
  // Create MCP connections if configured
  let mcpConnections: MCPConnection[] = [];
  if (body.options.mcp?.servers && body.options.mcp.servers.length > 0) {
    mcpConnections = await createMCPConnections(
      body.options.mcp.servers as MCPServerConfig[],
    );
  }

  // Prepare options with MCP connections
  const optionsWithMCP: z.infer<typeof callOptionsWithMCPSchema> = {
    ...body.options,
    mcpConnections,
  };

  const stream = createAgentUIStreamBase<
    z.infer<typeof callOptionsWithMCPSchema>,
    typeof tools,
    never,
    z.infer<typeof messageMetadataSchema>
  >({
    agent: SpreadsheetAgent,
    options: optionsWithMCP,
    sendSources: true,
    uiMessages: body.messages,
    experimental_transform: [smoothStream()],
    messageMetadata: ({ part }) => {
      if (part.type === "finish") {
        return { model: body.options.model, ...part.totalUsage };
      }
    },
    onFinish: async () => {
      // Close MCP connections when the stream finishes
      if (mcpConnections.length > 0) {
        await closeMCPConnections(mcpConnections);
      }
    },
  });

  return stream;
}

export async function chatRoute(req: Request) {
  const body = (await req.json()) as {
    messages: SpreadsheetAgentUIMessage[];
    options: z.infer<typeof callOptionsSchema>;
  };

  if (!body.options.anthropicApiKey) {
    return new Response("API key is required", { status: 400 });
  }

  const stream = await createAgentUIStream({ body });
  return createUIMessageStreamResponse({ stream });
}
