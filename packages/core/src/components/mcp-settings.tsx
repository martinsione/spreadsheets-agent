"use client";

import { PlusIcon, ServerIcon, Trash2Icon } from "lucide-react";
import { useState } from "react";
import type * as z from "zod";
import type { mcpServerConfigSchema } from "../ai/schema";
import { cn } from "../lib/utils";
import { Button } from "./ui/button";
import {
  Dialog,
  DialogContent,
  DialogDescription,
  DialogHeader,
  DialogTitle,
  DialogTrigger,
} from "./ui/dialog";
import { Input } from "./ui/input";
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from "./ui/select";

export type MCPServerConfig = z.infer<typeof mcpServerConfigSchema>;

interface MCPSettingsProps {
  servers: MCPServerConfig[];
  onServersChange: (servers: MCPServerConfig[]) => void;
}

interface MCPServerFormData {
  id: string;
  name: string;
  transportType: "sse" | "http";
  url: string;
}

function generateId() {
  return `mcp-${Date.now()}-${Math.random().toString(36).slice(2, 9)}`;
}

function MCPServerForm({
  initialData,
  onSubmit,
  onCancel,
  submitLabel,
}: {
  initialData?: MCPServerFormData;
  onSubmit: (data: MCPServerFormData) => void;
  onCancel: () => void;
  submitLabel: string;
}) {
  const [name, setName] = useState(initialData?.name ?? "");
  const [url, setUrl] = useState(initialData?.url ?? "");
  const [transportType, setTransportType] = useState<"sse" | "http">(
    initialData?.transportType ?? "sse",
  );
  const [errors, setErrors] = useState<{ name?: string; url?: string }>({});

  const validate = () => {
    const newErrors: { name?: string; url?: string } = {};
    if (!name.trim()) {
      newErrors.name = "Name is required";
    }
    if (!url.trim()) {
      newErrors.url = "URL is required";
    } else {
      try {
        new URL(url);
      } catch {
        newErrors.url = "Invalid URL format";
      }
    }
    setErrors(newErrors);
    return Object.keys(newErrors).length === 0;
  };

  const handleSubmit = () => {
    if (validate()) {
      onSubmit({
        id: initialData?.id ?? generateId(),
        name: name.trim(),
        url: url.trim(),
        transportType,
      });
    }
  };

  return (
    <div className="space-y-4">
      <div className="space-y-2">
        <label htmlFor="mcp-name" className="font-medium text-sm">
          Server Name
        </label>
        <Input
          id="mcp-name"
          placeholder="My MCP Server"
          value={name}
          onChange={(e) => setName(e.target.value)}
          aria-invalid={!!errors.name}
        />
        {errors.name && (
          <p className="text-destructive text-xs">{errors.name}</p>
        )}
      </div>

      <div className="space-y-2">
        <label htmlFor="mcp-transport" className="font-medium text-sm">
          Transport Type
        </label>
        <Select
          value={transportType}
          onValueChange={(v) => setTransportType(v as "sse" | "http")}
        >
          <SelectTrigger className="w-full">
            <SelectValue />
          </SelectTrigger>
          <SelectContent>
            <SelectItem value="sse">SSE (Server-Sent Events)</SelectItem>
            <SelectItem value="http">HTTP</SelectItem>
          </SelectContent>
        </Select>
        <p className="text-muted-foreground text-xs">
          SSE is recommended for most MCP servers.
        </p>
      </div>

      <div className="space-y-2">
        <label htmlFor="mcp-url" className="font-medium text-sm">
          Server URL
        </label>
        <Input
          id="mcp-url"
          type="url"
          placeholder="https://my-mcp-server.com/sse"
          value={url}
          onChange={(e) => setUrl(e.target.value)}
          aria-invalid={!!errors.url}
        />
        {errors.url && <p className="text-destructive text-xs">{errors.url}</p>}
      </div>

      <div className="flex justify-end gap-2 pt-2">
        <Button variant="outline" onClick={onCancel}>
          Cancel
        </Button>
        <Button onClick={handleSubmit}>{submitLabel}</Button>
      </div>
    </div>
  );
}

function MCPServerItem({
  server,
  onEdit,
  onDelete,
}: {
  server: MCPServerConfig;
  onEdit: () => void;
  onDelete: () => void;
}) {
  return (
    <div className="flex items-center justify-between rounded-md border border-border bg-muted/30 px-3 py-2">
      <div className="flex items-center gap-2 overflow-hidden">
        <ServerIcon className="size-4 shrink-0 text-muted-foreground" />
        <div className="min-w-0">
          <p className="truncate font-medium text-sm">{server.name}</p>
          <p className="truncate text-muted-foreground text-xs">
            {server.transport.type.toUpperCase()} • {server.transport.url}
          </p>
        </div>
      </div>
      <div className="flex shrink-0 items-center gap-1">
        <Button variant="ghost" size="sm" onClick={onEdit}>
          Edit
        </Button>
        <Button variant="ghost" size="icon-sm" onClick={onDelete}>
          <Trash2Icon className="size-3.5 text-destructive" />
        </Button>
      </div>
    </div>
  );
}

export function MCPSettings({ servers, onServersChange }: MCPSettingsProps) {
  const [isAddDialogOpen, setIsAddDialogOpen] = useState(false);
  const [editingServer, setEditingServer] = useState<MCPServerConfig | null>(
    null,
  );

  const handleAddServer = (data: MCPServerFormData) => {
    const newServer: MCPServerConfig = {
      id: data.id,
      name: data.name,
      transport: {
        type: data.transportType,
        url: data.url,
      },
    };
    onServersChange([...servers, newServer]);
    setIsAddDialogOpen(false);
  };

  const handleEditServer = (data: MCPServerFormData) => {
    const updatedServers = servers.map((s) =>
      s.id === data.id
        ? {
            ...s,
            name: data.name,
            transport: {
              type: data.transportType,
              url: data.url,
            },
          }
        : s,
    );
    onServersChange(updatedServers);
    setEditingServer(null);
  };

  const handleDeleteServer = (id: string) => {
    onServersChange(servers.filter((s) => s.id !== id));
  };

  return (
    <div className="space-y-3">
      <div className="flex items-center justify-between">
        <span className="font-medium text-sm">MCP Servers</span>
        <Dialog open={isAddDialogOpen} onOpenChange={setIsAddDialogOpen}>
          <DialogTrigger asChild>
            <Button variant="outline" size="sm">
              <PlusIcon className="size-3.5" />
              Add Server
            </Button>
          </DialogTrigger>
          <DialogContent>
            <DialogHeader>
              <DialogTitle>Add MCP Server</DialogTitle>
              <DialogDescription>
                Connect to an MCP server to access additional tools.
              </DialogDescription>
            </DialogHeader>
            <MCPServerForm
              onSubmit={handleAddServer}
              onCancel={() => setIsAddDialogOpen(false)}
              submitLabel="Add Server"
            />
          </DialogContent>
        </Dialog>
      </div>

      {servers.length === 0 ? (
        <div className="rounded-md border border-border border-dashed px-4 py-6 text-center">
          <ServerIcon className="mx-auto size-8 text-muted-foreground/50" />
          <p className="mt-2 text-muted-foreground text-sm">
            No MCP servers configured
          </p>
          <p className="text-muted-foreground text-xs">
            Add a server to access external tools
          </p>
        </div>
      ) : (
        <div className="space-y-2">
          {servers.map((server) => (
            <MCPServerItem
              key={server.id}
              server={server}
              onEdit={() => setEditingServer(server)}
              onDelete={() => handleDeleteServer(server.id)}
            />
          ))}
        </div>
      )}

      {/* Edit Dialog */}
      <Dialog
        open={editingServer !== null}
        onOpenChange={(open) => !open && setEditingServer(null)}
      >
        <DialogContent>
          <DialogHeader>
            <DialogTitle>Edit MCP Server</DialogTitle>
            <DialogDescription>
              Update the MCP server configuration.
            </DialogDescription>
          </DialogHeader>
          {editingServer && (
            <MCPServerForm
              initialData={{
                id: editingServer.id,
                name: editingServer.name,
                transportType: editingServer.transport.type,
                url: editingServer.transport.url,
              }}
              onSubmit={handleEditServer}
              onCancel={() => setEditingServer(null)}
              submitLabel="Save Changes"
            />
          )}
        </DialogContent>
      </Dialog>
    </div>
  );
}

/**
 * A simpler trigger button that can be used to open MCP settings.
 */
export function MCPSettingsTrigger({
  serverCount,
  className,
  ...props
}: React.ComponentProps<typeof Button> & { serverCount: number }) {
  return (
    <Button
      variant="outline"
      size="sm"
      className={cn("gap-1.5", className)}
      {...props}
    >
      <ServerIcon className="size-3.5" />
      MCP
      {serverCount > 0 && (
        <span className="rounded-full bg-primary px-1.5 py-0.5 font-medium text-primary-foreground text-xs leading-none">
          {serverCount}
        </span>
      )}
    </Button>
  );
}
