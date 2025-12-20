import { GASClient } from "gas-client";
import React from "react";
import { createRoot } from "react-dom/client";
import type { SpreadsheetService } from "@/spreadsheet-service";
import App from "./App";
import "./index.css";

function renderApp(spreadsheetService: SpreadsheetService) {
  const container = document.getElementById("root");
  if (container) {
    const root = createRoot(container);
    root.render(
      <React.StrictMode>
        <App
          spreadsheetService={spreadsheetService}
          environment="google-sheets"
        />
      </React.StrictMode>,
    );
  }
}

const { serverFunctions } = new GASClient<SpreadsheetService>({
  // this is necessary for local development but will be ignored in production
  allowedDevelopmentDomains: (origin) =>
    /https:\/\/.*\.googleusercontent\.com$/.test(origin),
});

renderApp(serverFunctions as unknown as SpreadsheetService);
