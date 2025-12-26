import type { MCPClient, MCPClientConfig, MCPTransport } from "@ai-sdk/mcp";
import { createMCPClient } from "@ai-sdk/mcp";
import type { Tool } from "ai";

/**
 * MCP transport configuration for SSE or HTTP transports.
 */
export interface MCPTransportConfig {
  type: "sse" | "http";
  url: string;
  headers?: Record<string, string>;
}

/**
 * Configuration for a single MCP server.
 */
export interface MCPServerConfig {
  /** Unique identifier for this MCP server */
  id: string;
  /** Human-readable name for the server */
  name: string;
  /** Transport configuration */
  transport: MCPTransportConfig | MCPTransport;
  /** Optional: Tool names to require approval before execution */
  toolsRequiringApproval?: string[];
}

/**
 * Holds an MCP client and its associated tools.
 */
export interface MCPConnection {
  client: MCPClient;
  tools: Record<string, Tool>;
  config: MCPServerConfig;
}

/**
 * Creates an MCP client connection and fetches its tools.
 */
export async function createMCPConnection(
  serverConfig: MCPServerConfig,
): Promise<MCPConnection> {
  const clientConfig: MCPClientConfig = {
    transport: serverConfig.transport,
    name: `opensheets-mcp-${serverConfig.id}`,
  };

  const client = await createMCPClient(clientConfig);
  const tools = await client.tools();

  return {
    client,
    tools,
    config: serverConfig,
  };
}

/**
 * Creates multiple MCP connections in parallel.
 */
export async function createMCPConnections(
  serverConfigs: MCPServerConfig[],
): Promise<MCPConnection[]> {
  const connections = await Promise.all(
    serverConfigs.map((config) => createMCPConnection(config)),
  );
  return connections;
}

/**
 * Merges MCP tools with existing tools, prefixing MCP tool names with server ID
 * to avoid conflicts.
 */
export function mergeMCPTools(
  existingTools: Record<string, Tool>,
  mcpConnections: MCPConnection[],
  options?: {
    /** If true, prefix MCP tool names with server ID (default: false) */
    prefixWithServerId?: boolean;
    /** Tool names that should require approval */
    toolsRequiringApproval?: string[];
  },
): Record<string, Tool> {
  const mergedTools = { ...existingTools };
  const { prefixWithServerId = false, toolsRequiringApproval = [] } =
    options ?? {};

  for (const connection of mcpConnections) {
    const serverApprovalTools = connection.config.toolsRequiringApproval ?? [];
    const allApprovalTools = [
      ...toolsRequiringApproval,
      ...serverApprovalTools,
    ];

    for (const [toolName, tool] of Object.entries(connection.tools)) {
      const finalToolName = prefixWithServerId
        ? `${connection.config.id}_${toolName}`
        : toolName;

      // Check if this tool requires approval
      const needsApproval =
        allApprovalTools.includes(toolName) ||
        allApprovalTools.includes(finalToolName);

      mergedTools[finalToolName] = needsApproval
        ? { ...tool, needsApproval: true }
        : tool;
    }
  }

  return mergedTools;
}

/**
 * Closes all MCP connections gracefully.
 */
export async function closeMCPConnections(
  connections: MCPConnection[],
): Promise<void> {
  await Promise.all(connections.map((conn) => conn.client.close()));
}
