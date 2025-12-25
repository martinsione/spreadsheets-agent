"use client";

import "@mescius/spread-sheets/styles/gc.spread.sheets.excel2016colorful.css";
import "@mescius/spread-sheets-designer/styles/gc.spread.sheets.designer.min.css";
import "@mescius/spread-sheets-designer-resources-en";
import "@mescius/spread-sheets-charts";
import "@mescius/spread-sheets-pivot-addon";
import GC from "@mescius/spread-sheets";
import * as GCDesigner from "@mescius/spread-sheets-designer";
import { Designer } from "@mescius/spread-sheets-designer-react";
import { useCallback, useRef } from "react";

// Set SpreadJS license keys
if (process.env.NEXT_PUBLIC_SPREADJS_LICENSE_KEY) {
  GC.Spread.Sheets.LicenseKey = process.env.NEXT_PUBLIC_SPREADJS_LICENSE_KEY;
}
if (process.env.NEXT_PUBLIC_SPREADJS_DESIGNER_LICENSE_KEY) {
  GCDesigner.Spread.Sheets.Designer.LicenseKey =
    process.env.NEXT_PUBLIC_SPREADJS_DESIGNER_LICENSE_KEY;
}

type SpreadsheetProps = {
  onInitialized?: (workbook: GC.Spread.Sheets.Workbook) => void;
};

export function Spreadsheet({ onInitialized }: SpreadsheetProps) {
  const workbookRef = useRef<GC.Spread.Sheets.Workbook | null>(null);

  const handleDesignerInit = useCallback(
    (designer: GCDesigner.Spread.Sheets.Designer.Designer) => {
      const spread = designer.getWorkbook() as GC.Spread.Sheets.Workbook;
      workbookRef.current = spread;

      spread.options.tabStripVisible = true;
      spread.options.allowUserDragDrop = true;
      spread.options.allowUserDragFill = true;
      spread.options.allowUserResize = true;
      spread.options.allowContextMenu = true;
      spread.options.allowUserEditFormula = true;

      const sheet = spread.getActiveSheet();
      if (sheet) {
        sheet.setRowCount(1000);
        sheet.setColumnCount(26);
      }
      onInitialized?.(spread);
    },
    [onInitialized],
  );

  return (
    <Designer
      designerInitialized={handleDesignerInit}
      styleInfo={{ width: "100%", height: "100%" }}
    />
  );
}
