// Types

// Client
export {
  connectToMCP,
  disconnectFromMCP,
  getActiveClient,
  getAllActiveClients,
  getAllMCPTools,
  handleOAuthCallback,
} from "./client";

// Storage
export {
  addConnection,
  clearOAuthTokens,
  deleteConnection,
  getConnection,
  loadAndClearOAuthState,
  loadConnections,
  saveConnections,
  saveOAuthState,
  saveOAuthTokens,
  toggleConnection,
  updateConnection,
} from "./storage";
export type {
  MCPAuthType,
  MCPConnectionConfig,
  MCPConnectionFormData,
  MCPConnectionState,
  MCPConnectionStatus,
  MCPConnectionsStorage,
  MCPOAuthState,
  MCPOAuthTokens,
  MCPToolInfo,
} from "./types";

// Hooks
export { type UseMCPReturn, useMCP } from "./use-mcp";
