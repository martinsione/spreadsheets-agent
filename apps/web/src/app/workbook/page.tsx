"use client";

import type GC from "@mescius/spread-sheets";
import { Chat } from "@repo/core/components/chat";
import type { SpreadsheetService } from "@repo/core/spreadsheet-service";
import dynamic from "next/dynamic";
import { useEffect, useMemo, useRef, useState } from "react";

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
  const [spreadsheetService, setSpreadsheetService] =
    useState<SpreadsheetService | null>(null);

  useEffect(() => {
    // Dynamically import the service to avoid SSR issues with SpreadJS
    import("@/lib/spreadsheet-service").then(
      ({ createWebSpreadsheetService }) => {
        setSpreadsheetService(
          createWebSpreadsheetService(() => workbookRef.current),
        );
      },
    );
  }, []);

  return (
    <div className="flex h-full w-full">
      <div className="contain-[paint] h-full flex-1 overflow-hidden">
        <Spreadsheet
          onInitialized={(workbook) => {
            workbookRef.current = workbook;
          }}
        />
      </div>

      <div className="h-full w-[420px] overflow-hidden border-border border-l bg-background">
        {spreadsheetService ? (
          <Chat spreadsheetService={spreadsheetService} environment="web" />
        ) : (
          <div className="flex h-full items-center justify-center text-muted-foreground">
            Loading...
          </div>
        )}
      </div>
    </div>
  );
}
