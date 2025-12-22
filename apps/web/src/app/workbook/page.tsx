"use client";

import type GC from "@mescius/spread-sheets";
import { Chat } from "@repo/core/components/chat";
import dynamic from "next/dynamic";
import { useMemo, useRef } from "react";
import { createWebSpreadsheetService } from "@/lib/spreadsheet-service";

const Spreadsheet = dynamic(
  () => import("@/components/spreadsheet").then((mod) => mod.Spreadsheet),
  {
    ssr: false,
    loading: () => (
      <div className="flex h-full w-full animate-pulse items-center justify-center text-muted-foreground duration-300">
        Loading spreadsheet...
      </div>
    ),
  },
);

export default function WorkbookPage() {
  const workbookRef = useRef<GC.Spread.Sheets.Workbook | null>(null);
  const spreadsheetService = useMemo(
    () => createWebSpreadsheetService(() => workbookRef.current),
    [],
  );

  return (
    <div className="flex h-full w-full">
      <div className="h-full flex-1 overflow-hidden">
        <Spreadsheet
          onInitialized={(workbook) => {
            workbookRef.current = workbook;
          }}
        />
      </div>

      <div className="h-full w-[300px] overflow-hidden border-border border-l bg-background">
        <Chat spreadsheetService={spreadsheetService} environment="web" />
      </div>
    </div>
  );
}
