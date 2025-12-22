import { Chat } from "@repo/core/components/chat";
import type { SpreadsheetService } from "@repo/core/spreadsheet-service";
import React from "react";
import { createRoot } from "react-dom/client";
import { excelService } from "@/spreadsheet-service";
import "@repo/core/styles.css";

function renderApp(spreadsheetService: SpreadsheetService) {
  const container = document.getElementById("root");
  if (container) {
    const root = createRoot(container);
    root.render(
      <React.StrictMode>
        <Chat spreadsheetService={spreadsheetService} environment="excel" />
      </React.StrictMode>,
    );
  }
}

renderApp(excelService);
