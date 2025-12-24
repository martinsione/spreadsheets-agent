"use client";

import { PlugIcon, PlusIcon } from "lucide-react";
import { useState } from "react";
import { getConnection } from "../../mcp/storage";
import type {
  MCPConnectionConfig,
  MCPConnectionFormData,
} from "../../mcp/types";
import type { UseMCPReturn } from "../../mcp/use-mcp";
import { Button } from "../ui/button";
import {
  Dialog,
  DialogContent,
  DialogDescription,
  DialogHeader,
  DialogTitle,
  DialogTrigger,
} from "../ui/dialog";
import { MCPConnectionDialog } from "./mcp-connection-dialog";
import { MCPConnectionList } from "./mcp-connection-list";

interface MCPManagerProps {
  mcp: UseMCPReturn;
}

export function MCPManager({ mcp }: MCPManagerProps) {
  const [managerOpen, setManagerOpen] = useState(false);
  const [addDialogOpen, setAddDialogOpen] = useState(false);
  const [editDialogOpen, setEditDialogOpen] = useState(false);
  const [editingConnection, setEditingConnection] =
    useState<MCPConnectionConfig | null>(null);

  const connectedCount = mcp.connections.filter(
    (c) => c.status === "connected",
  ).length;

  const handleAdd = async (formData: MCPConnectionFormData) => {
    await mcp.addConnection(formData);
  };

  const handleEdit = (id: string) => {
    const connection = getConnection(id);
    if (connection) {
      setEditingConnection(connection);
      setEditDialogOpen(true);
    }
  };

  const handleUpdate = async (formData: MCPConnectionFormData) => {
    if (editingConnection) {
      await mcp.updateConnection(editingConnection.id, formData);
      setEditingConnection(null);
    }
  };

  return (
    <>
      <Dialog open={managerOpen} onOpenChange={setManagerOpen}>
        <DialogTrigger asChild>
          <Button variant="ghost" size="icon" className="relative">
            <PlugIcon className="size-4" />
            {connectedCount > 0 && (
              <span className="absolute -top-0.5 -right-0.5 flex size-3 items-center justify-center rounded-full bg-green-500 font-medium text-[0.5rem] text-white">
                {connectedCount}
              </span>
            )}
          </Button>
        </DialogTrigger>
        <DialogContent className="sm:max-w-md">
          <DialogHeader>
            <DialogTitle>MCP Connections</DialogTitle>
            <DialogDescription>
              Manage your Model Context Protocol connections to extend AI
              capabilities with custom tools.
            </DialogDescription>
          </DialogHeader>

          <MCPConnectionList
            connections={mcp.connections}
            onEdit={handleEdit}
            onDelete={mcp.deleteConnection}
            onConnect={mcp.connect}
            onDisconnect={mcp.disconnect}
          />

          <Button
            onClick={() => setAddDialogOpen(true)}
            className="w-full"
            variant="outline"
          >
            <PlusIcon className="size-3" />
            Add MCP Connection
          </Button>
        </DialogContent>
      </Dialog>

      <MCPConnectionDialog
        open={addDialogOpen}
        onOpenChange={setAddDialogOpen}
        onSubmit={handleAdd}
        mode="add"
      />

      <MCPConnectionDialog
        open={editDialogOpen}
        onOpenChange={(open) => {
          setEditDialogOpen(open);
          if (!open) setEditingConnection(null);
        }}
        onSubmit={handleUpdate}
        initialData={editingConnection ?? undefined}
        mode="edit"
      />
    </>
  );
}
