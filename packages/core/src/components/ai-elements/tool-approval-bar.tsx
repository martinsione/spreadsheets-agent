"use client";

import { Hand } from "lucide-react";
import { useCallback, useEffectEvent } from "react";
import type { tools } from "../../ai/tools";
import { Button } from "../ui/button";

export type ToolApprovalBarProps = {
  toolName: keyof typeof tools;
  explanation: string;
  onApprove: () => void;
  onApproveAll: () => void;
  onDecline: () => void;
};

export function ToolApprovalBar({
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

  return (
    <div className="absolute bottom-0 z-10 w-full border-t bg-muted px-3 py-6">
      <div className="mb-2 flex items-center gap-2">
        <Hand className="size-4 shrink-0 text-muted-foreground" />
        <span className="font-medium text-sm">Permission required</span>
      </div>
      <p className="mb-3 text-muted-foreground text-sm">{explanation}</p>
      <div className="flex flex-col gap-2">
        <Button type="button" onClick={onDecline} variant="outline">
          Decline
          <kbd className="rounded bg-muted px-1.5 py-0.5 font-mono text-muted-foreground text-xs">
            ESC
          </kbd>
        </Button>

        <Button type="button" onClick={onApproveAll} variant="outline">
          Allow all
          <kbd className="rounded bg-muted px-1.5 py-0.5 font-mono text-muted-foreground text-xs">
            ⇧↵
          </kbd>
        </Button>

        <Button type="button" variant="default" onClick={onApprove}>
          Allow
          <kbd className="rounded bg-primary-foreground/20 px-1.5 py-0.5 font-mono text-xs">
            ↵
          </kbd>
        </Button>
      </div>
    </div>
  );
}
