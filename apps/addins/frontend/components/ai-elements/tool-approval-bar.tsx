"use client";

import { Hand } from "lucide-react";
import { useCallback, useEffectEvent } from "react";
import type { tools, writeTools } from "@/server/ai/tools";
import { Button } from "../ui/button";

export type ToolApprovalBarProps = {
  toolName: keyof typeof tools;
  explanation: string;
  onApprove: () => void;
  onApproveAll: () => void;
  onDecline: () => void;
};

function formatToolName(toolName: (typeof writeTools)[number]): string {
  const nameMap: Record<(typeof writeTools)[number], string> = {
    clearCellRange: "clear cells",
    copyTo: "copy data",
    modifyObject: "modify object",
    modifySheetStructure: "modify sheet structure",
    modifyWorkbookStructure: "modify workbook",
    setCellRange: "edit cells",
  };

  return nameMap[toolName] || toolName.replace(/([A-Z])/g, " $1").toLowerCase();
}

export function ToolApprovalBar({
  toolName,
  explanation,
  onApprove,
  onApproveAll,
  onDecline,
}: ToolApprovalBarProps) {
  const handleKeyDown = useCallback(
    (e: KeyboardEvent) => {
      if (e.key === "Escape") {
        e.preventDefault();
        onDecline();
      } else if (e.shiftKey && e.key === "Enter") {
        e.preventDefault();
        onApproveAll();
      } else if (e.key === "Enter") {
        e.preventDefault();
        onApprove();
      }
    },
    [onDecline, onApprove, onApproveAll],
  );

  useEffectEvent(() => {
    document.addEventListener("keydown", handleKeyDown);
    return () => document.removeEventListener("keydown", handleKeyDown);
  });

  const toolDisplayName = toolName
    ? formatToolName(toolName as (typeof writeTools)[number])
    : "perform action";

  return (
    <div className="fixed bottom-0 left-1/2 z-50 w-full max-w-4xl -translate-x-1/2 rounded-t-3xl border border-t bg-background px-6 py-4">
      <div className="flex flex-col gap-2 pb-4">
        <div className="flex items-center gap-2">
          <Hand className="size-4 text-muted-foreground" />
          <span className="font-medium text-sm">Permission required</span>
        </div>
        <span className="text-muted-foreground text-sm">
          Agent wants to {explanation}
        </span>
      </div>

      <div className="flex flex-col gap-2">
        <Button type="button" variant="default" onClick={onApprove}>
          Allow
          <kbd className="rounded bg-primary-foreground/20 px-1.5 py-0.5 font-mono text-xs">
            ↵
          </kbd>
        </Button>

        <Button type="button" onClick={onApproveAll} variant="outline">
          Allow all edits
          <kbd className="rounded bg-muted px-1.5 py-0.5 font-mono text-muted-foreground text-xs"></kbd>
          <kbd className="rounded bg-muted px-1.5 py-0.5 font-mono text-muted-foreground text-xs">
            ⇧ + ↵
          </kbd>
        </Button>

        <Button type="button" onClick={onDecline} variant="outline">
          Decline
          <kbd className="rounded bg-muted px-1.5 py-0.5 font-mono text-muted-foreground text-xs">
            ESC
          </kbd>
        </Button>
      </div>
    </div>
  );
}
