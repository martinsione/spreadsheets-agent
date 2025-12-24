"use client";

import {
  KeyIcon,
  LinkIcon,
  LockIcon,
  PlusIcon,
  UnlockIcon,
  XIcon,
} from "lucide-react";
import { useState } from "react";
import { cn } from "../../lib/utils";
import type {
  MCPAuthType,
  MCPConnectionConfig,
  MCPConnectionFormData,
} from "../../mcp/types";
import { Button } from "../ui/button";
import { ButtonGroup } from "../ui/button-group";
import {
  Dialog,
  DialogContent,
  DialogDescription,
  DialogFooter,
  DialogHeader,
  DialogTitle,
} from "../ui/dialog";
import { Input } from "../ui/input";

interface MCPConnectionDialogProps {
  open: boolean;
  onOpenChange: (open: boolean) => void;
  onSubmit: (data: MCPConnectionFormData) => Promise<void>;
  initialData?: MCPConnectionConfig;
  mode: "add" | "edit";
}

const AUTH_TYPES: {
  value: MCPAuthType;
  label: string;
  icon: React.ReactNode;
}[] = [
  { value: "none", label: "None", icon: <UnlockIcon className="size-3" /> },
  { value: "bearer", label: "Bearer", icon: <KeyIcon className="size-3" /> },
  { value: "headers", label: "Headers", icon: <LinkIcon className="size-3" /> },
  { value: "oauth", label: "OAuth", icon: <LockIcon className="size-3" /> },
];

export function MCPConnectionDialog({
  open,
  onOpenChange,
  onSubmit,
  initialData,
  mode,
}: MCPConnectionDialogProps) {
  const [name, setName] = useState(initialData?.name || "");
  const [url, setUrl] = useState(initialData?.url || "");
  const [authType, setAuthType] = useState<MCPAuthType>(
    initialData?.authType || "none",
  );
  const [bearerToken, setBearerToken] = useState(
    initialData?.bearerToken || "",
  );
  const [headers, setHeaders] = useState<Array<{ key: string; value: string }>>(
    initialData?.headers
      ? Object.entries(initialData.headers).map(([key, value]) => ({
          key,
          value: String(value),
        }))
      : [{ key: "", value: "" }],
  );
  const [isSubmitting, setIsSubmitting] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    setError(null);

    if (!name.trim()) {
      setError("Name is required");
      return;
    }

    if (!url.trim()) {
      setError("URL is required");
      return;
    }

    try {
      new URL(url);
    } catch {
      setError("Invalid URL format");
      return;
    }

    setIsSubmitting(true);

    try {
      const formData: MCPConnectionFormData = {
        name: name.trim(),
        url: url.trim(),
        authType,
        ...(authType === "bearer" && { bearerToken }),
        ...(authType === "headers" && {
          headers: Object.fromEntries(
            headers
              .filter((h) => h.key.trim() && h.value.trim())
              .map((h) => [h.key.trim(), h.value.trim()]),
          ),
        }),
      };

      await onSubmit(formData);
      onOpenChange(false);

      // Reset form
      setName("");
      setUrl("");
      setAuthType("none");
      setBearerToken("");
      setHeaders([{ key: "", value: "" }]);
    } catch (err) {
      setError(
        err instanceof Error ? err.message : "Failed to save connection",
      );
    } finally {
      setIsSubmitting(false);
    }
  };

  const addHeader = () => {
    setHeaders([...headers, { key: "", value: "" }]);
  };

  const removeHeader = (index: number) => {
    setHeaders(headers.filter((_, i) => i !== index));
  };

  const updateHeader = (
    index: number,
    field: "key" | "value",
    value: string,
  ) => {
    setHeaders(
      headers.map((h, i) => (i === index ? { ...h, [field]: value } : h)),
    );
  };

  return (
    <Dialog open={open} onOpenChange={onOpenChange}>
      <DialogContent className="sm:max-w-md">
        <DialogHeader>
          <DialogTitle>
            {mode === "add" ? "Add MCP Connection" : "Edit MCP Connection"}
          </DialogTitle>
          <DialogDescription>
            Connect to a Model Context Protocol server to extend AI capabilities
            with custom tools.
          </DialogDescription>
        </DialogHeader>

        <form onSubmit={handleSubmit} className="space-y-4">
          <div className="space-y-2">
            <label htmlFor="name" className="font-medium text-xs">
              Name
            </label>
            <Input
              id="name"
              placeholder="My MCP Server"
              value={name}
              onChange={(e) => setName(e.target.value)}
            />
          </div>

          <div className="space-y-2">
            <label htmlFor="url" className="font-medium text-xs">
              URL
            </label>
            <Input
              id="url"
              type="url"
              placeholder="https://mcp.example.com"
              value={url}
              onChange={(e) => setUrl(e.target.value)}
            />
          </div>

          <fieldset className="space-y-2">
            <legend className="font-medium text-xs">Authentication</legend>
            <ButtonGroup className="w-full">
              {AUTH_TYPES.map((type) => (
                <Button
                  key={type.value}
                  type="button"
                  variant={authType === type.value ? "default" : "outline"}
                  className="flex-1"
                  onClick={() => setAuthType(type.value)}
                >
                  {type.icon}
                  {type.label}
                </Button>
              ))}
            </ButtonGroup>
          </fieldset>

          {authType === "bearer" && (
            <div className="space-y-2">
              <label htmlFor="bearer" className="font-medium text-xs">
                Bearer Token
              </label>
              <Input
                id="bearer"
                type="password"
                placeholder="Enter your bearer token"
                value={bearerToken}
                onChange={(e) => setBearerToken(e.target.value)}
              />
            </div>
          )}

          {authType === "headers" && (
            <fieldset className="space-y-2">
              <legend className="font-medium text-xs">Custom Headers</legend>
              <div className="space-y-2">
                {headers.map((header, index) => (
                  <div key={index} className="flex gap-2">
                    <Input
                      placeholder="Header name"
                      value={header.key}
                      onChange={(e) =>
                        updateHeader(index, "key", e.target.value)
                      }
                      className="flex-1"
                    />
                    <Input
                      placeholder="Value"
                      value={header.value}
                      onChange={(e) =>
                        updateHeader(index, "value", e.target.value)
                      }
                      className="flex-1"
                    />
                    {headers.length > 1 && (
                      <Button
                        type="button"
                        variant="ghost"
                        size="icon"
                        onClick={() => removeHeader(index)}
                      >
                        <XIcon className="size-3" />
                      </Button>
                    )}
                  </div>
                ))}
                <Button
                  type="button"
                  variant="outline"
                  size="sm"
                  onClick={addHeader}
                  className="w-full"
                >
                  <PlusIcon className="size-3" />
                  Add Header
                </Button>
              </div>
            </fieldset>
          )}

          {authType === "oauth" && (
            <div className="rounded-md border border-border bg-muted/50 p-3 text-muted-foreground text-xs">
              You'll be redirected to authorize after saving this MCP
              connection.
            </div>
          )}

          {error && (
            <div className="rounded-md border border-destructive/50 bg-destructive/10 p-2 text-destructive text-xs">
              {error}
            </div>
          )}

          <DialogFooter>
            <Button
              type="button"
              variant="outline"
              onClick={() => onOpenChange(false)}
            >
              Cancel
            </Button>
            <Button type="submit" disabled={isSubmitting}>
              {isSubmitting
                ? "Saving..."
                : mode === "add"
                  ? authType === "oauth"
                    ? "Authorize"
                    : "Add"
                  : "Save"}
            </Button>
          </DialogFooter>
        </form>
      </DialogContent>
    </Dialog>
  );
}
