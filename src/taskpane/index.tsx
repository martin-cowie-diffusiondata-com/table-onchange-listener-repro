/* eslint-disable @typescript-eslint/no-unused-vars */
import React from "react";
import { FluentProvider, webLightTheme } from "@fluentui/react-components";
import { createRoot } from "react-dom/client";

import { App } from "./components/App";

// eslint-disable-next-line no-redeclare
/* global document, Office */
let isOfficeInitialized = false;

// Render application after Office initializes
Office.onReady(() => {
  isOfficeInitialized = true;
  createRoot(document.getElementById("container") as HTMLElement).render(
    <React.StrictMode>
      <FluentProvider theme={webLightTheme}>
        <App />
      </FluentProvider>
    </React.StrictMode>
  );
});
