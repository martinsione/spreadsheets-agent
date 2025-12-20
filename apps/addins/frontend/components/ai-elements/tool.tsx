"use client";

import type { ToolUIPart } from "ai";
import {
  CheckCircleIcon,
  ChevronDownIcon,
  CircleIcon,
  ClockIcon,
  WrenchIcon,
  XCircleIcon,
} from "lucide-react";
import type { ComponentProps, ReactNode } from "react";
import { isValidElement } from "react";
import {
  CellRangePreview,
  type CellRangePreviewProps,
} from "@/frontend/components/ai-elements/cell-range-preview";
import { CodeBlock } from "@/frontend/components/ai-elements/code-block";
import { Badge } from "@/frontend/components/ui/badge";
import {
  Collapsible,
  CollapsibleContent,
  CollapsibleTrigger,
} from "@/frontend/components/ui/collapsible";
import { cn } from "@/frontend/lib/utils";

export type ToolProps = ComponentProps<typeof Collapsible>;

export const Tool = ({ className, ...props }: ToolProps) => (
  <Collapsible
    className={cn("not-prose mb-4 w-full rounded-md border", className)}
    {...props}
  />
);

export type ToolHeaderProps = {
  title?: string;
  type: ToolUIPart["type"];
  state: ToolUIPart["state"];
  className?: string;
};

const getStatusBadge = (status: ToolUIPart["state"]) => {
  const labels: Record<ToolUIPart["state"], string> = {
    "input-streaming": "Pending",
    "input-available": "Running",
    "approval-requested": "Awaiting Approval",
    "approval-responded": "Responded",
    "output-available": "Completed",
    "output-error": "Error",
    "output-denied": "Denied",
  };

  const icons: Record<ToolUIPart["state"], ReactNode> = {
    "input-streaming": <CircleIcon className="size-4" />,
    "input-available": <ClockIcon className="size-4 animate-pulse" />,
    "approval-requested": <ClockIcon className="size-4 text-yellow-600" />,
    "approval-responded": <CheckCircleIcon className="size-4 text-blue-600" />,
    "output-available": <CheckCircleIcon className="size-4 text-green-600" />,
    "output-error": <XCircleIcon className="size-4 text-red-600" />,
    "output-denied": <XCircleIcon className="size-4 text-orange-600" />,
  };

  return (
    <Badge className="gap-1.5 rounded-full text-xs" variant="secondary">
      {icons[status]}
      {labels[status]}
    </Badge>
  );
};

export const ToolHeader = ({
  className,
  title,
  type,
  state,
  ...props
}: ToolHeaderProps) => (
  <CollapsibleTrigger
    className={cn(
      "flex w-full items-center justify-between gap-4 p-3",
      className,
    )}
    {...props}
  >
    <div className="flex items-center gap-2 truncate">
      <WrenchIcon className="size-4 shrink-0 text-muted-foreground" />
      <span
        className={cn(
          "truncate font-medium text-sm",
          state === "input-streaming"
            ? "animate-pulse text-muted-foreground"
            : "",
        )}
      >
        {title}
      </span>
      {/* {getStatusBadge(state)} */}
    </div>
    <ChevronDownIcon className="size-4 text-muted-foreground transition-transform group-data-[state=open]:rotate-180" />
  </CollapsibleTrigger>
);

export type ToolContentProps = ComponentProps<typeof CollapsibleContent>;

export const ToolContent = ({ className, ...props }: ToolContentProps) => (
  <CollapsibleContent
    className={cn(
      "data-[state=closed]:fade-out-0 data-[state=closed]:slide-out-to-top-2 data-[state=open]:slide-in-from-top-2 text-popover-foreground outline-none data-[state=closed]:animate-out data-[state=open]:animate-in",
      className,
    )}
    {...props}
  />
);

export type ToolInputProps = ComponentProps<"div"> & {
  toolName?: string;
  input: ToolUIPart["input"];
};

export const ToolInput = ({
  className,
  toolName,
  input,
  ...props
}: ToolInputProps) => {
  const typedInput = input as Record<string, unknown> | undefined;

  if (
    toolName === "setCellRange" &&
    typedInput?.range &&
    Array.isArray(typedInput?.cells)
  ) {
    return (
      <div
        className={cn("space-y-2 overflow-hidden p-4", className)}
        {...props}
      >
        <h4 className="font-medium text-muted-foreground text-xs uppercase tracking-wide">
          Parameters
        </h4>
        <CellRangePreview
          range={typedInput.range as string}
          cells={typedInput.cells as CellRangePreviewProps["cells"]}
          explanation={typedInput.explanation as string | undefined}
        />
      </div>
    );
  }

  return (
    <div className={cn("space-y-2 overflow-hidden p-4", className)} {...props}>
      <h4 className="font-medium text-muted-foreground text-xs uppercase tracking-wide">
        Parameters
      </h4>
      <div className="rounded-md bg-muted/50">
        <CodeBlock code={JSON.stringify(input, null, 2)} language="json" />
      </div>
    </div>
  );
};

export type ToolOutputProps = ComponentProps<"div"> & {
  toolName?: string;
  state: ToolUIPart["state"];
  output: ToolUIPart["output"];
  errorText: ToolUIPart["errorText"];
};

/**
 * Transforms getCellRanges output format into CellRangePreview format.
 * Input: { cells: { "A1": "value", "B2": [computedValue, "=formula"] } }
 * Output: 2D array of Cell objects matching the grid layout
 */
function transformCellRangesOutput(
  dimension: string,
  cells: Record<string, unknown>,
): { range: string; cells: CellRangePreviewProps["cells"] } | null {
  const rangeMatch = dimension.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/i);
  if (!rangeMatch) return null;

  const [, startColStr, startRowStr, endColStr, endRowStr] = rangeMatch;
  if (!startColStr || !startRowStr || !endColStr || !endRowStr) return null;

  const colToIndex = (col: string): number => {
    let index = 0;
    for (let i = 0; i < col.length; i++) {
      index = index * 26 + (col.toUpperCase().charCodeAt(i) - 64);
    }
    return index - 1;
  };

  const indexToCol = (index: number): string => {
    let column = "";
    let num = index + 1;
    while (num > 0) {
      const remainder = (num - 1) % 26;
      column = String.fromCharCode(65 + remainder) + column;
      num = Math.floor((num - 1) / 26);
    }
    return column;
  };

  const startCol = colToIndex(startColStr);
  const endCol = colToIndex(endColStr);
  const startRow = parseInt(startRowStr, 10);
  const endRow = parseInt(endRowStr, 10);

  const numRows = endRow - startRow + 1;
  const numCols = endCol - startCol + 1;

  const grid: CellRangePreviewProps["cells"] = [];

  for (let r = 0; r < numRows; r++) {
    const row: CellRangePreviewProps["cells"][number] = [];
    for (let c = 0; c < numCols; c++) {
      const cellAddress = `${indexToCol(startCol + c)}${startRow + r}`;
      const cellData = cells[cellAddress];

      if (cellData === undefined || cellData === null) {
        row.push({ value: "" });
      } else if (Array.isArray(cellData)) {
        // Format: [computedValue, formula, optionalNote?]
        const [value, formula, note] = cellData;
        row.push({
          value: value ?? "",
          formula: typeof formula === "string" ? formula : undefined,
          note: typeof note === "string" ? note : undefined,
        });
      } else {
        // Simple value
        row.push({
          value: cellData as string | number | boolean,
        });
      }
    }
    grid.push(row);
  }

  return { range: dimension, cells: grid };
}

export const ToolOutput = ({
  className,
  toolName,
  state,
  output,
  errorText,
  ...props
}: ToolOutputProps) => {
  if (state !== "output-available" && state !== "output-error") {
    return null;
  }

  if (!(output || errorText)) {
    return null;
  }

  // Handle getCellRanges output with grid preview
  if (toolName === "getCellRanges" && typeof output === "object" && output) {
    const typedOutput = output as {
      worksheet?: {
        name?: string;
        dimension?: string;
        cells?: Record<string, unknown>;
      };
      hasMore?: boolean;
    };

    if (
      typedOutput.worksheet?.dimension &&
      typedOutput.worksheet?.cells &&
      typeof typedOutput.worksheet.cells === "object"
    ) {
      const transformed = transformCellRangesOutput(
        typedOutput.worksheet.dimension,
        typedOutput.worksheet.cells,
      );

      if (transformed) {
        return (
          <div className={cn("space-y-2 p-4", className)} {...props}>
            <h4 className="font-medium text-muted-foreground text-xs uppercase tracking-wide">
              Result
            </h4>
            <div className="space-y-2">
              <div className="flex items-center gap-2 text-muted-foreground text-xs">
                <span className="font-medium">
                  {typedOutput.worksheet.name}
                </span>
                <span>•</span>
                <span>{typedOutput.worksheet.dimension}</span>
                {typedOutput.hasMore && (
                  <>
                    <span>•</span>
                    <span className="text-yellow-600">
                      Truncated (more data available)
                    </span>
                  </>
                )}
              </div>
              <CellRangePreview
                range={transformed.range}
                cells={transformed.cells}
              />
            </div>
          </div>
        );
      }
    }
  }

  let Output = <div>{output as ReactNode}</div>;

  if (typeof output === "object" && !isValidElement(output)) {
    Output = (
      <CodeBlock code={JSON.stringify(output, null, 2)} language="json" />
    );
  } else if (typeof output === "string") {
    Output = <CodeBlock code={output} language="json" />;
  }

  return (
    <div className={cn("space-y-2 p-4", className)} {...props}>
      <h4 className="font-medium text-muted-foreground text-xs uppercase tracking-wide">
        {errorText ? "Error" : "Result"}
      </h4>
      <div
        className={cn(
          "overflow-x-auto rounded-md text-xs [&_table]:w-full",
          errorText
            ? "bg-destructive/10 text-destructive"
            : "bg-muted/50 text-foreground",
        )}
      >
        {errorText && <div>{errorText}</div>}
        {Output}
      </div>
    </div>
  );
};
