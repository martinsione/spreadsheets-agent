"use client";

import type GC from "@mescius/spread-sheets";
import { Chat } from "@repo/core/components/chat";
import {
  Tabs,
  TabsContent,
  TabsList,
  TabsTrigger,
} from "@repo/core/components/ui/tabs";
import { MessageSquareIcon, TableIcon } from "lucide-react";
import { useMemo, useRef } from "react";
import { Spreadsheet } from "@/components/spreadsheet";
import { createWebSpreadsheetService } from "@/lib/spreadsheet-service";

export function WorkbookContent() {
  const workbookRef = useRef<GC.Spread.Sheets.Workbook | null>(null);

  const spreadsheetService = useMemo(() => {
    return createWebSpreadsheetService(() => workbookRef.current);
  }, []);

  const handleSpreadsheetInit = (workbook: GC.Spread.Sheets.Workbook) => {
    workbookRef.current = workbook;
  };

  return (
    <>
      <Tabs
        defaultValue="chat"
        className="flex h-full w-full flex-col md:hidden"
      >
        <TabsList className="w-full shrink-0 gap-0 rounded-none border-border border-b bg-muted/50 p-0">
          <TabsTrigger
            value="spreadsheet"
            className="py-3.5 text-sm transition-all data-[state=active]:text-foreground"
          >
            <TableIcon className="size-4" />
            <span>Spreadsheet</span>
          </TabsTrigger>
          <TabsTrigger
            value="chat"
            className="py-3.5 text-sm transition-all data-[state=active]:text-foreground"
          >
            <MessageSquareIcon className="size-4" />
            <span>Chat</span>
          </TabsTrigger>
        </TabsList>

        <TabsContent value="spreadsheet" className="mt-0 h-full flex-1">
          <div className="contain-[paint] h-full w-full overflow-hidden">
            <Spreadsheet onInitialized={handleSpreadsheetInit} />
          </div>
        </TabsContent>

        <TabsContent value="chat" className="mt-0 h-full flex-1">
          <Chat spreadsheetService={spreadsheetService} environment="web" />
        </TabsContent>
      </Tabs>

      {/* Desktop view with side-by-side layout (md and above) */}
      <div className="hidden h-full w-full md:flex">
        <div className="h-full flex-1">
          <div className="contain-[paint] h-full w-full overflow-hidden">
            <Spreadsheet onInitialized={handleSpreadsheetInit} />
          </div>
        </div>

        <div className="h-full w-[420px] shrink-0 border-border border-l">
          <Chat spreadsheetService={spreadsheetService} environment="web" />
        </div>
      </div>
    </>
  );
}
