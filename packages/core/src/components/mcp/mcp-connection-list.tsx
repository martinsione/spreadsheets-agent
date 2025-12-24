"use client";

import {
  ChevronDownIcon,
  CircleIcon,
  EditIcon,
  PlugIcon,
  PowerIcon,
  RefreshCwIcon,
  TrashIcon,
  WrenchIcon,
} from "lucide-react";
import { useState } from "react";
import { cn } from "../../lib/utils";
import type {
  MCPConnectionFormData,
  MCPConnectionState,
} from "../../mcp/types";
import { Badge } from "../ui/badge";
import { Button } from "../ui/button";
import {
  Collapsible,
  CollapsibleContent,
  CollapsibleTrigger,
} from "../ui/collapsible";

interface MCPConnectionListProps {
  connections: MCPConnectionState[];
  onEdit: (id: string) => void;
  onDelete: (id: string) => Promise<void>;
  onConnect: (id: string) => Promise<void>;
  onDisconnect: (id: string) => Promise<void>;
}

const STATUS_COLORS: Record<MCPConnectionState["status"], string> = {
  connected: "bg-green-500",
  connecting: "bg-yellow-500 animate-pulse",
  disconnected: "bg-gray-400",
  error: "bg-red-500",
  "oauth-pending": "bg-blue-500 animate-pulse",
};

const STATUS_LABELS: Record<MCPConnectionState["status"], string> = {
  connected: "Connected",
  connecting: "Connecting...",
  disconnected: "Disconnected",
  error: "Error",
  "oauth-pending": "Awaiting OAuth",
};

export function MCPConnectionList({
  connections,
  onEdit,
  onDelete,
  onConnect,
  onDisconnect,
}: MCPConnectionListProps) {
  const [expandedId, setExpandedId] = useState<string | null>(null);
  const [deletingId, setDeletingId] = useState<string | null>(null);

  if (connections.length === 0) {
    return (
      <div className="flex flex-col items-center justify-center gap-2 py-8 text-center">
        <PlugIcon className="size-8 text-muted-foreground/50" />
        <p className="text-muted-foreground text-xs">
          No MCP connections configured.
        </p>
        <p className="text-muted-foreground/70 text-xs">
          Add a connection to extend AI capabilities with custom tools.
        </p>
      </div>
    );
  }

  return (
    <div className="space-y-2">
      {connections.map((connection) => (
        <Collapsible
          key={connection.config.id}
          open={expandedId === connection.config.id}
          onOpenChange={(open) =>
            setExpandedId(open ? connection.config.id : null)
          }
        >
          <div className="rounded-lg border border-border bg-card">
            <div className="flex items-center gap-2 p-3">
              <div
                className={cn(
                  "size-2 rounded-full",
                  STATUS_COLORS[connection.status],
                )}
                title={STATUS_LABELS[connection.status]}
              />

              <div className="min-w-0 flex-1">
                <div className="flex items-center gap-2">
                  <span className="truncate font-medium text-xs">
                    {connection.config.name}
                  </span>
                  <Badge variant="outline" className="text-[0.5rem]">
                    {connection.config.authType}
                  </Badge>
                </div>
                <p className="truncate text-[0.625rem] text-muted-foreground">
                  {connection.config.url}
                </p>
              </div>

              <div className="flex items-center gap-1">
                {connection.status === "connected" &&
                  connection.tools.length > 0 && (
                    <CollapsibleTrigger asChild>
                      <Button variant="ghost" size="icon-sm">
                        <ChevronDownIcon
                          className={cn(
                            "size-3 transition-transform",
                            expandedId === connection.config.id && "rotate-180",
                          )}
                        />
                      </Button>
                    </CollapsibleTrigger>
                  )}

                {connection.status === "connected" ? (
                  <Button
                    variant="ghost"
                    size="icon-sm"
                    onClick={() => onDisconnect(connection.config.id)}
                    title="Disconnect"
                  >
                    <PowerIcon className="size-3" />
                  </Button>
                ) : connection.status === "connecting" ? (
                  <Button variant="ghost" size="icon-sm" disabled>
                    <RefreshCwIcon className="size-3 animate-spin" />
                  </Button>
                ) : (
                  <Button
                    variant="ghost"
                    size="icon-sm"
                    onClick={() => onConnect(connection.config.id)}
                    title="Connect"
                  >
                    <PlugIcon className="size-3" />
                  </Button>
                )}

                <Button
                  variant="ghost"
                  size="icon-sm"
                  onClick={() => onEdit(connection.config.id)}
                  title="Edit"
                >
                  <EditIcon className="size-3" />
                </Button>

                <Button
                  variant="ghost"
                  size="icon-sm"
                  onClick={async () => {
                    setDeletingId(connection.config.id);
                    await onDelete(connection.config.id);
                    setDeletingId(null);
                  }}
                  disabled={deletingId === connection.config.id}
                  title="Delete"
                >
                  <TrashIcon className="size-3" />
                </Button>
              </div>
            </div>

            {connection.error && (
              <div className="border-border border-t px-3 py-2">
                <p className="text-[0.625rem] text-destructive">
                  {connection.error}
                </p>
              </div>
            )}

            <CollapsibleContent>
              {connection.tools.length > 0 && (
                <div className="border-border border-t p-3">
                  <div className="mb-2 flex items-center gap-1 text-muted-foreground">
                    <WrenchIcon className="size-3" />
                    <span className="font-medium text-[0.625rem]">
                      Available Tools ({connection.tools.length})
                    </span>
                  </div>
                  <div className="flex flex-wrap gap-1">
                    {connection.tools.map((tool) => (
                      <Badge
                        key={tool.name}
                        variant="secondary"
                        className="text-[0.5rem]"
                        title={tool.description}
                      >
                        {tool.name}
                      </Badge>
                    ))}
                  </div>
                </div>
              )}
            </CollapsibleContent>
          </div>
        </Collapsible>
      ))}
    </div>
  );
}
