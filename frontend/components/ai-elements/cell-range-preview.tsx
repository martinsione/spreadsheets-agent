"use client";

import type { CSSProperties } from "react";
import { cn } from "@/frontend/lib/utils";

type CellStyles = {
  backgroundColor?: string;
  fontColor?: string;
  fontSize?: number;
  fontWeight?: "normal" | "bold";
  fontStyle?: "normal" | "italic";
  textDecoration?: "none" | "underline" | "line-through";
  horizontalAlignment?: "left" | "center" | "right";
  verticalAlignment?: "top" | "middle" | "bottom";
};

type Cell = {
  value?: string | number | boolean;
  formula?: string;
  note?: string;
  cellStyles?: CellStyles;
};

export type CellRangePreviewProps = {
  range: string;
  cells: Cell[][];
  explanation?: string;
  className?: string;
};

function parseRange(range: string): {
  startCol: string;
  startRow: number;
  endCol: string;
  endRow: number;
} {
  const defaultRange = { startCol: "A", startRow: 1, endCol: "A", endRow: 1 };

  const rangeMatch = range.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/i);
  if (rangeMatch) {
    const [, startCol, startRow, endCol, endRow] = rangeMatch;
    if (startCol && startRow && endCol && endRow) {
      return {
        startCol: startCol.toUpperCase(),
        startRow: parseInt(startRow, 10),
        endCol: endCol.toUpperCase(),
        endRow: parseInt(endRow, 10),
      };
    }
  }

  const singleMatch = range.match(/^([A-Z]+)(\d+)$/i);
  if (singleMatch) {
    const [, col, row] = singleMatch;
    if (col && row) {
      return {
        startCol: col.toUpperCase(),
        startRow: parseInt(row, 10),
        endCol: col.toUpperCase(),
        endRow: parseInt(row, 10),
      };
    }
  }

  return defaultRange;
}

function columnToIndex(col: string): number {
  let index = 0;
  for (let i = 0; i < col.length; i++) {
    index = index * 26 + (col.charCodeAt(i) - 64);
  }
  return index - 1;
}

function indexToColumn(index: number): string {
  let column = "";
  let num = index + 1;
  while (num > 0) {
    const remainder = (num - 1) % 26;
    column = String.fromCharCode(65 + remainder) + column;
    num = Math.floor((num - 1) / 26);
  }
  return column;
}

function getColumnHeaders(startCol: string, endCol: string): string[] {
  const startIndex = columnToIndex(startCol);
  const endIndex = columnToIndex(endCol);
  const headers: string[] = [];
  for (let i = startIndex; i <= endIndex; i++) {
    headers.push(indexToColumn(i));
  }
  return headers;
}

function cellStylesToCSS(styles?: CellStyles): CSSProperties {
  if (!styles) return {};

  const css: CSSProperties = {};

  if (styles.backgroundColor) {
    css.backgroundColor = styles.backgroundColor;
  }
  if (styles.fontColor) {
    css.color = styles.fontColor;
  }
  if (styles.fontSize) {
    css.fontSize = `${styles.fontSize}px`;
  }
  if (styles.fontWeight) {
    css.fontWeight = styles.fontWeight;
  }
  if (styles.fontStyle) {
    css.fontStyle = styles.fontStyle;
  }
  if (styles.textDecoration) {
    css.textDecoration = styles.textDecoration;
  }
  if (styles.horizontalAlignment) {
    css.textAlign = styles.horizontalAlignment;
  }
  if (styles.verticalAlignment) {
    css.verticalAlign = styles.verticalAlignment;
  }

  return css;
}

function formatCellValue(cell: Cell): string {
  if (cell.formula) {
    return cell.formula;
  }
  if (cell.value === undefined || cell.value === null) {
    return "";
  }
  if (typeof cell.value === "boolean") {
    return cell.value ? "TRUE" : "FALSE";
  }
  return String(cell.value);
}

export function CellRangePreview({
  range,
  cells,
  explanation,
  className,
}: CellRangePreviewProps) {
  const { startCol, startRow } = parseRange(range);
  const columnHeaders = getColumnHeaders(
    startCol,
    indexToColumn(columnToIndex(startCol) + (cells[0]?.length || 1) - 1),
  );

  return (
    <div className={cn("space-y-2", className)}>
      {explanation && (
        <p className="text-muted-foreground text-xs">{explanation}</p>
      )}
      <div className="overflow-x-auto rounded-md border">
        <table className="w-full border-collapse text-xs">
          <thead>
            <tr className="bg-muted/50">
              <th className="w-10 border-r border-b bg-muted/80 px-2 py-1 text-center font-medium text-muted-foreground" />
              {columnHeaders.map((col) => (
                <th
                  key={col}
                  className="min-w-[60px] border-r border-b bg-muted/80 px-2 py-1 text-center font-medium text-muted-foreground last:border-r-0"
                >
                  {col}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {cells.map((row, rowIndex) => (
              <tr key={rowIndex} className="border-b last:border-b-0">
                <td className="border-r bg-muted/50 px-2 py-1 text-center font-medium text-muted-foreground">
                  {startRow + rowIndex}
                </td>
                {row.map((cell, colIndex) => {
                  const displayValue = formatCellValue(cell);
                  const cellStyle = cellStylesToCSS(cell.cellStyles);

                  return (
                    <td
                      key={colIndex}
                      className="max-w-[150px] truncate border-r px-2 py-1 last:border-r-0"
                      style={cellStyle}
                      title={
                        displayValue.length > 20 ? displayValue : undefined
                      }
                    >
                      {displayValue}
                    </td>
                  );
                })}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}
