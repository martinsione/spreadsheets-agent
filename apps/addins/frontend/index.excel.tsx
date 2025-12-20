import React from "react";
import { createRoot } from "react-dom/client";
import type { SpreadsheetService } from "@/spreadsheet-service";
import { excelService } from "@/spreadsheet-service/excel";
import App from "./App";
import "./index.css";

function renderApp(spreadsheetService: SpreadsheetService) {
  const container = document.getElementById("root");
  if (container) {
    const root = createRoot(container);
    root.render(
      <React.StrictMode>
        <App spreadsheetService={spreadsheetService} environment="excel" />
      </React.StrictMode>,
    );
  }
}

renderApp(excelService);
