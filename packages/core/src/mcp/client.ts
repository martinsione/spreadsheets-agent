import {
  auth,
  createMCPClient,
  type MCPClient,
  type OAuthClientProvider,
} from "@ai-sdk/mcp";
import { nanoid } from "nanoid";
import {
  clearOAuthTokens,
  loadAndClearOAuthState,
  saveOAuthState,
  saveOAuthTokens,
} from "./storage";
import type {
  MCPConnectionConfig,
  MCPConnectionState,
  MCPConnectionStatus,
  MCPOAuthTokens,
  MCPToolInfo,
} from "./types";

/**
 * Create an OAuth client provider for MCP connections
 */
function createOAuthProvider(
  connection: MCPConnectionConfig,
  callbacks: {
    onTokensSaved: (tokens: MCPOAuthTokens) => void;
    onRedirect: (url: URL) => void;
  },
): OAuthClientProvider {
  let codeVerifier: string | undefined;
  let currentTokens: MCPOAuthTokens | undefined = connection.oauthTokens;

  return {
    get redirectUrl(): string {
      // Use the current origin for redirect
      if (typeof window !== "undefined") {
        return `${window.location.origin}${window.location.pathname}`;
      }
      return "https://opensheets.app/workbook";
    },

    get clientMetadata() {
      return {
        redirect_uris: [this.redirectUrl] as string[],
        client_name: "OpenSheets",
        client_uri: "https://opensheets.app",
      };
    },

    tokens() {
      if (!currentTokens) return undefined;
      return {
        access_token: currentTokens.access_token,
        token_type: currentTokens.token_type || "Bearer",
        expires_in: currentTokens.expires_in,
        refresh_token: currentTokens.refresh_token,
        scope: currentTokens.scope,
      };
    },

    async saveTokens(tokens: MCPOAuthTokens) {
      currentTokens = tokens;
      callbacks.onTokensSaved(tokens);
    },

    async redirectToAuthorization(authorizationUrl: URL) {
      // Save state for OAuth callback
      saveOAuthState({
        codeVerifier: codeVerifier || "",
        state: authorizationUrl.searchParams.get("state") || "",
        connectionId: connection.id,
      });
      callbacks.onRedirect(authorizationUrl);
    },

    async saveCodeVerifier(verifier: string) {
      codeVerifier = verifier;
    },

    codeVerifier() {
      return codeVerifier || "";
    },

    clientInformation() {
      return undefined;
    },
  };
}

/**
 * Active MCP clients cache
 */
const activeClients = new Map<string, MCPClient>();

/**
 * Connect to an MCP server
 */
export async function connectToMCP(
  connection: MCPConnectionConfig,
  callbacks: {
    onStatusChange: (status: MCPConnectionStatus, error?: string) => void;
    onToolsLoaded: (tools: MCPToolInfo[]) => void;
    onTokensSaved: (tokens: MCPOAuthTokens) => void;
    onOAuthRedirect: (url: URL) => void;
  },
): Promise<MCPClient | null> {
  // Close existing client if any
  await disconnectFromMCP(connection.id);

  callbacks.onStatusChange("connecting");

  try {
    // Build headers based on auth type
    const headers: Record<string, string> = {};

    if (connection.authType === "bearer" && connection.bearerToken) {
      headers.Authorization = `Bearer ${connection.bearerToken}`;
    } else if (connection.authType === "headers" && connection.headers) {
      Object.assign(headers, connection.headers);
    }

    // Create client config
    const clientConfig: Parameters<typeof createMCPClient>[0] = {
      transport: {
        type: "sse",
        url: connection.url,
        headers: Object.keys(headers).length > 0 ? headers : undefined,
        ...(connection.authType === "oauth"
          ? {
              authProvider: createOAuthProvider(connection, {
                onTokensSaved: callbacks.onTokensSaved,
                onRedirect: callbacks.onOAuthRedirect,
              }),
            }
          : {}),
      },
      name: `opensheets-mcp-${connection.id}`,
      onUncaughtError: (error) => {
        console.error(`MCP ${connection.name} error:`, error);
        callbacks.onStatusChange(
          "error",
          error instanceof Error ? error.message : "Unknown error",
        );
      },
    };

    const client = await createMCPClient(clientConfig);
    activeClients.set(connection.id, client);

    // Fetch available tools
    const toolsResult = await client.tools();
    const tools: MCPToolInfo[] = Object.entries(toolsResult).map(
      ([name, tool]) => ({
        name,
        description: (tool as { description?: string }).description,
        inputSchema: (
          tool as { parameters?: { jsonSchema?: Record<string, unknown> } }
        ).parameters?.jsonSchema,
      }),
    );

    callbacks.onToolsLoaded(tools);
    callbacks.onStatusChange("connected");

    return client;
  } catch (error) {
    console.error(`Failed to connect to MCP ${connection.name}:`, error);

    // Check if this is an OAuth redirect needed
    if (error instanceof Error && error.message.includes("Unauthorized")) {
      callbacks.onStatusChange("oauth-pending", "OAuth authorization required");
      return null;
    }

    callbacks.onStatusChange(
      "error",
      error instanceof Error ? error.message : "Connection failed",
    );
    return null;
  }
}

/**
 * Disconnect from an MCP server
 */
export async function disconnectFromMCP(connectionId: string): Promise<void> {
  const client = activeClients.get(connectionId);
  if (client) {
    try {
      await client.close();
    } catch (error) {
      console.error(`Error closing MCP client ${connectionId}:`, error);
    }
    activeClients.delete(connectionId);
  }
}

/**
 * Get an active MCP client
 */
export function getActiveClient(connectionId: string): MCPClient | undefined {
  return activeClients.get(connectionId);
}

/**
 * Get all active clients
 */
export function getAllActiveClients(): Map<string, MCPClient> {
  return new Map(activeClients);
}

/**
 * Handle OAuth callback
 */
export async function handleOAuthCallback(
  _code: string,
  state: string,
): Promise<{ connectionId: string } | null> {
  const savedState = loadAndClearOAuthState();

  if (!savedState || savedState.state !== state) {
    console.error("OAuth state mismatch");
    return null;
  }

  // The actual token exchange will happen when we reconnect
  // The auth provider will have the code verifier saved
  return { connectionId: savedState.connectionId };
}

/**
 * Get all tools from all connected MCPs
 */
export async function getAllMCPTools(): Promise<
  Record<
    string,
    ReturnType<MCPClient["tools"]> extends Promise<infer T> ? T : never
  >
> {
  const result: Record<string, Awaited<ReturnType<MCPClient["tools"]>>> = {};

  for (const [connectionId, client] of activeClients) {
    try {
      const tools = await client.tools();
      result[connectionId] = tools;
    } catch (error) {
      console.error(`Failed to get tools from ${connectionId}:`, error);
    }
  }

  return result;
}
