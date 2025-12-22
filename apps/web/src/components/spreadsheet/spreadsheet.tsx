"use client";

import "@mescius/spread-sheets/styles/gc.spread.sheets.excel2016colorful.css";
import "@mescius/spread-sheets-charts";
import "@mescius/spread-sheets-pivot-addon";
import type GC from "@mescius/spread-sheets";
import { SpreadSheets } from "@mescius/spread-sheets-react";
import { useCallback, useRef } from "react";

export interface SpreadsheetHandle {
  getWorkbook: () => GC.Spread.Sheets.Workbook | null;
}

interface SpreadsheetProps {
  onInitialized?: (workbook: GC.Spread.Sheets.Workbook) => void;
  className?: string;
}

export function Spreadsheet({ onInitialized, className }: SpreadsheetProps) {
  const workbookRef = useRef<GC.Spread.Sheets.Workbook | null>(null);

  const handleWorkbookInit = useCallback(
    (spread: GC.Spread.Sheets.Workbook) => {
      workbookRef.current = spread;

      spread.options.tabStripVisible = true;
      spread.options.allowUserDragDrop = true;
      spread.options.allowUserDragFill = true;
      spread.options.allowUserResize = true;
      spread.options.allowContextMenu = true;
      spread.options.allowUserEditFormula = true;

      let retries = 0;
      const waitForReady = () => {
        if (spread.getSheetCount() > 0) {
          const sheet = spread.getActiveSheet();
          if (sheet) {
            sheet.setRowCount(1000);
            sheet.setColumnCount(26);
          }
          onInitialized?.(spread);
        } else if (retries++ < 100) {
          requestAnimationFrame(waitForReady);
        } else {
          spread.addSheet(0);
          const sheet = spread.getSheet(0);
          if (sheet) {
            sheet.name("Sheet1");
            sheet.setRowCount(1000);
            sheet.setColumnCount(26);
          }
          onInitialized?.(spread);
        }
      };
      requestAnimationFrame(waitForReady);
    },
    [onInitialized],
  );

  return (
    <SpreadSheets
      workbookInitialized={handleWorkbookInit}
      hostStyle={{ width: "100%", height: "100%" }}
      hostClass={className}
    />
  );
}
