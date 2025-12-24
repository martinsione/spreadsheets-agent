"use client";

import type { MCPClient } from "@ai-sdk/mcp";
import { useCallback, useEffect, useRef, useState } from "react";
import { connectToMCP, disconnectFromMCP, getAllActiveClients } from "./client";
import {
  addConnection as addConnectionToStorage,
  deleteConnection as deleteConnectionFromStorage,
  getConnection,
  loadConnections,
  saveOAuthTokens,
  updateConnection as updateConnectionInStorage,
} from "./storage";
import type {
  MCPConnectionConfig,
  MCPConnectionFormData,
  MCPConnectionState,
  MCPConnectionStatus,
  MCPToolInfo,
} from "./types";

export interface UseMCPReturn {
  /** All connection states */
  connections: MCPConnectionState[];
  /** Add a new MCP connection */
  addConnection: (
    formData: MCPConnectionFormData,
  ) => Promise<MCPConnectionConfig>;
  /** Update an existing connection */
  updateConnection: (
    id: string,
    updates: Partial<MCPConnectionFormData>,
  ) => Promise<MCPConnectionConfig | null>;
  /** Delete a connection */
  deleteConnection: (id: string) => Promise<void>;
  /** Connect to an MCP server */
  connect: (id: string) => Promise<void>;
  /** Disconnect from an MCP server */
  disconnect: (id: string) => Promise<void>;
  /** Reconnect all enabled connections */
  reconnectAll: () => Promise<void>;
  /** Get all active MCP clients */
  getClients: () => Map<string, MCPClient>;
  /** Get combined tools from all connected MCPs */
  getAllTools: () => Record<string, MCPToolInfo[]>;
  /** Loading state */
  isLoading: boolean;
}

export function useMCP(): UseMCPReturn {
  const [connections, setConnections] = useState<MCPConnectionState[]>([]);
  const [isLoading, setIsLoading] = useState(true);
  const initRef = useRef(false);

  const updateConnectionState = useCallback(
    (
      id: string,
      updates: Partial<Omit<MCPConnectionState, "config">> & {
        config?: MCPConnectionConfig;
      },
    ) => {
      setConnections((prev) =>
        prev.map((conn) =>
          conn.config.id === id ? { ...conn, ...updates } : conn,
        ),
      );
    },
    [],
  );

  const connectToConnection = useCallback(
    async (id: string) => {
      const config = getConnection(id);
      if (!config) return;

      await connectToMCP(config, {
        onStatusChange: (status, error) => {
          updateConnectionState(id, { status, error });
        },
        onToolsLoaded: (tools) => {
          updateConnectionState(id, { tools });
        },
        onTokensSaved: (tokens) => {
          const updated = saveOAuthTokens(id, tokens);
          if (updated) {
            updateConnectionState(id, { config: updated });
          }
        },
        onOAuthRedirect: (url) => {
          // Open OAuth URL in new window or redirect
          window.open(url.toString(), "_blank", "width=600,height=700");
        },
      });
    },
    [updateConnectionState],
  );

  // Initialize connections from storage
  useEffect(() => {
    if (initRef.current) return;
    initRef.current = true;

    const storedConnections = loadConnections();
    const initialStates: MCPConnectionState[] = storedConnections.map(
      (config) => ({
        config,
        status: "disconnected" as MCPConnectionStatus,
        tools: [],
      }),
    );
    setConnections(initialStates);
    setIsLoading(false);

    // Auto-connect enabled connections
    for (const state of initialStates) {
      if (state.config.enabled) {
        connectToConnection(state.config.id);
      }
    }
  }, [connectToConnection]);

  const addConnection = useCallback(
    async (formData: MCPConnectionFormData): Promise<MCPConnectionConfig> => {
      const newConfig = addConnectionToStorage(formData);

      const newState: MCPConnectionState = {
        config: newConfig,
        status: "disconnected",
        tools: [],
      };

      setConnections((prev) => [...prev, newState]);

      // Auto-connect if enabled
      if (newConfig.enabled) {
        await connectToConnection(newConfig.id);
      }

      return newConfig;
    },
    [connectToConnection],
  );

  const updateConnection = useCallback(
    async (
      id: string,
      updates: Partial<MCPConnectionFormData>,
    ): Promise<MCPConnectionConfig | null> => {
      const updated = updateConnectionInStorage(id, updates);
      if (!updated) return null;

      updateConnectionState(id, { config: updated });

      // Reconnect if URL or auth changed
      const needsReconnect =
        "url" in updates ||
        "authType" in updates ||
        "bearerToken" in updates ||
        "headers" in updates;

      if (needsReconnect && updated.enabled) {
        await disconnectFromMCP(id);
        await connectToConnection(id);
      }

      return updated;
    },
    [connectToConnection, updateConnectionState],
  );

  const deleteConnection = useCallback(async (id: string): Promise<void> => {
    await disconnectFromMCP(id);
    deleteConnectionFromStorage(id);
    setConnections((prev) => prev.filter((conn) => conn.config.id !== id));
  }, []);

  const connect = useCallback(
    async (id: string): Promise<void> => {
      const config = getConnection(id);
      if (!config) return;

      // Enable the connection in storage
      updateConnectionInStorage(id, { enabled: true });
      updateConnectionState(id, {
        config: { ...config, enabled: true },
      });

      await connectToConnection(id);
    },
    [connectToConnection, updateConnectionState],
  );

  const disconnect = useCallback(
    async (id: string): Promise<void> => {
      await disconnectFromMCP(id);

      const config = getConnection(id);
      if (config) {
        updateConnectionInStorage(id, { enabled: false });
        updateConnectionState(id, {
          config: { ...config, enabled: false },
          status: "disconnected",
          tools: [],
        });
      }
    },
    [updateConnectionState],
  );

  const reconnectAll = useCallback(async (): Promise<void> => {
    for (const conn of connections) {
      if (conn.config.enabled) {
        await disconnectFromMCP(conn.config.id);
        await connectToConnection(conn.config.id);
      }
    }
  }, [connections, connectToConnection]);

  const getClients = useCallback((): Map<string, MCPClient> => {
    return getAllActiveClients();
  }, []);

  const getAllTools = useCallback((): Record<string, MCPToolInfo[]> => {
    const result: Record<string, MCPToolInfo[]> = {};
    for (const conn of connections) {
      if (conn.status === "connected" && conn.tools.length > 0) {
        result[conn.config.id] = conn.tools;
      }
    }
    return result;
  }, [connections]);

  return {
    connections,
    addConnection,
    updateConnection,
    deleteConnection,
    connect,
    disconnect,
    reconnectAll,
    getClients,
    getAllTools,
    isLoading,
  };
}
