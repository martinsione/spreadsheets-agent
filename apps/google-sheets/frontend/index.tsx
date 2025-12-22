import { Chat } from "@repo/core/components/chat";
import type { SpreadsheetService } from "@repo/core/spreadsheet-service";
import { GASClient } from "gas-client";
import React from "react";
import { createRoot } from "react-dom/client";
import "@repo/core/styles.css";

function renderApp(spreadsheetService: SpreadsheetService) {
  const container = document.getElementById("root");
  if (container) {
    const root = createRoot(container);
    root.render(
      <React.StrictMode>
        <Chat
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
