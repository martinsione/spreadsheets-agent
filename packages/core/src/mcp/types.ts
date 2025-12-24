import * as z from "zod";

/**
 * Authentication types for MCP connections
 */
export const MCPAuthType = z.enum(["none", "bearer", "headers", "oauth"]);
export type MCPAuthType = z.infer<typeof MCPAuthType>;

/**
 * OAuth tokens for MCP connections
 */
export const MCPOAuthTokens = z.object({
  access_token: z.string(),
  token_type: z.string().optional(),
  expires_in: z.number().optional(),
  refresh_token: z.string().optional(),
  scope: z.string().optional(),
  expires_at: z.number().optional(),
});
export type MCPOAuthTokens = z.infer<typeof MCPOAuthTokens>;

/**
 * OAuth state stored for pending OAuth flows
 */
export const MCPOAuthState = z.object({
  codeVerifier: z.string(),
  state: z.string(),
  connectionId: z.string(),
});
export type MCPOAuthState = z.infer<typeof MCPOAuthState>;

/**
 * Configuration for a single MCP connection
 */
export const MCPConnectionConfig = z.object({
  /** Unique identifier for this connection */
  id: z.string(),
  /** Display name for this connection */
  name: z.string(),
  /** URL of the MCP server */
  url: z.string().url(),
  /** Authentication type */
  authType: MCPAuthType,
  /** Bearer token (when authType is "bearer") */
  bearerToken: z.string().optional(),
  /** Custom headers (when authType is "headers") */
  headers: z.record(z.string(), z.string()).optional(),
  /** OAuth tokens (when authType is "oauth") */
  oauthTokens: MCPOAuthTokens.optional(),
  /** Whether this connection is enabled */
  enabled: z.boolean(),
  /** Timestamp of when this connection was created */
  createdAt: z.number(),
  /** Timestamp of when this connection was last updated */
  updatedAt: z.number(),
});
export type MCPConnectionConfig = z.infer<typeof MCPConnectionConfig>;

/**
 * Connection status
 */
export type MCPConnectionStatus =
  | "disconnected"
  | "connecting"
  | "connected"
  | "error"
  | "oauth-pending";

/**
 * Tool information from an MCP server
 */
export interface MCPToolInfo {
  name: string;
  description?: string;
  inputSchema?: Record<string, unknown>;
}

/**
 * Runtime state for an MCP connection
 */
export interface MCPConnectionState {
  config: MCPConnectionConfig;
  status: MCPConnectionStatus;
  error?: string;
  tools: MCPToolInfo[];
}

/**
 * Stored connections configuration
 */
export const MCPConnectionsStorage = z.object({
  connections: z.array(MCPConnectionConfig),
  version: z.number().default(1),
});
export type MCPConnectionsStorage = z.infer<typeof MCPConnectionsStorage>;

/**
 * Form data for creating/editing an MCP connection
 */
export interface MCPConnectionFormData {
  name: string;
  url: string;
  authType: MCPAuthType;
  bearerToken?: string;
  headers?: Record<string, string>;
}
