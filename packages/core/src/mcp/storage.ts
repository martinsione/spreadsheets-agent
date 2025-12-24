import { nanoid } from "nanoid";
import type {
  MCPConnectionConfig,
  MCPConnectionFormData,
  MCPConnectionsStorage,
  MCPOAuthState,
  MCPOAuthTokens,
} from "./types";

const STORAGE_KEY = "opensheets_mcp_connections";
const OAUTH_STATE_KEY = "opensheets_mcp_oauth_state";

/**
 * Load MCP connections from localStorage
 */
export function loadConnections(): MCPConnectionConfig[] {
  if (typeof window === "undefined") return [];

  try {
    const stored = localStorage.getItem(STORAGE_KEY);
    if (!stored) return [];

    const parsed = JSON.parse(stored) as MCPConnectionsStorage;
    return parsed.connections || [];
  } catch {
    return [];
  }
}

/**
 * Save MCP connections to localStorage
 */
export function saveConnections(connections: MCPConnectionConfig[]): void {
  if (typeof window === "undefined") return;

  const storage: MCPConnectionsStorage = {
    connections,
    version: 1,
  };
  localStorage.setItem(STORAGE_KEY, JSON.stringify(storage));
}

/**
 * Add a new MCP connection
 */
export function addConnection(
  formData: MCPConnectionFormData,
): MCPConnectionConfig {
  const connections = loadConnections();
  const now = Date.now();

  const newConnection: MCPConnectionConfig = {
    id: nanoid(),
    name: formData.name,
    url: formData.url,
    authType: formData.authType,
    bearerToken: formData.bearerToken,
    headers: formData.headers,
    enabled: true,
    createdAt: now,
    updatedAt: now,
  };

  connections.push(newConnection);
  saveConnections(connections);

  return newConnection;
}

/**
 * Update an existing MCP connection
 */
export function updateConnection(
  id: string,
  updates: Partial<MCPConnectionFormData> & { enabled?: boolean },
): MCPConnectionConfig | null {
  const connections = loadConnections();
  const index = connections.findIndex((c) => c.id === id);

  if (index === -1) return null;

  const connection = connections[index];
  const updatedConnection: MCPConnectionConfig = {
    ...connection,
    ...updates,
    updatedAt: Date.now(),
  };

  connections[index] = updatedConnection;
  saveConnections(connections);

  return updatedConnection;
}

/**
 * Delete an MCP connection
 */
export function deleteConnection(id: string): boolean {
  const connections = loadConnections();
  const filtered = connections.filter((c) => c.id !== id);

  if (filtered.length === connections.length) return false;

  saveConnections(filtered);
  return true;
}

/**
 * Toggle connection enabled state
 */
export function toggleConnection(id: string): MCPConnectionConfig | null {
  const connections = loadConnections();
  const connection = connections.find((c) => c.id === id);

  if (!connection) return null;

  return updateConnection(id, {
    ...connection,
    enabled: !connection.enabled,
  } as MCPConnectionFormData);
}

/**
 * Get a single connection by ID
 */
export function getConnection(id: string): MCPConnectionConfig | null {
  const connections = loadConnections();
  return connections.find((c) => c.id === id) || null;
}

/**
 * Save OAuth state for a pending OAuth flow
 */
export function saveOAuthState(state: MCPOAuthState): void {
  if (typeof window === "undefined") return;
  localStorage.setItem(OAUTH_STATE_KEY, JSON.stringify(state));
}

/**
 * Load and clear OAuth state
 */
export function loadAndClearOAuthState(): MCPOAuthState | null {
  if (typeof window === "undefined") return null;

  try {
    const stored = localStorage.getItem(OAUTH_STATE_KEY);
    if (!stored) return null;

    localStorage.removeItem(OAUTH_STATE_KEY);
    return JSON.parse(stored) as MCPOAuthState;
  } catch {
    return null;
  }
}

/**
 * Save OAuth tokens for a connection
 */
export function saveOAuthTokens(
  connectionId: string,
  tokens: MCPOAuthTokens,
): MCPConnectionConfig | null {
  const connections = loadConnections();
  const index = connections.findIndex((c) => c.id === connectionId);

  if (index === -1) return null;

  const connection = connections[index];
  const updatedConnection: MCPConnectionConfig = {
    ...connection,
    oauthTokens: tokens,
    updatedAt: Date.now(),
  };

  connections[index] = updatedConnection;
  saveConnections(connections);

  return updatedConnection;
}

/**
 * Clear OAuth tokens for a connection
 */
export function clearOAuthTokens(connectionId: string): void {
  const connections = loadConnections();
  const index = connections.findIndex((c) => c.id === connectionId);

  if (index === -1) return;

  const connection = connections[index];
  const updatedConnection: MCPConnectionConfig = {
    ...connection,
    oauthTokens: undefined,
    updatedAt: Date.now(),
  };

  connections[index] = updatedConnection;
  saveConnections(connections);
}
