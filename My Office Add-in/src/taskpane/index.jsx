import * as React from "react";
import { createRoot } from "react-dom/client";
import App from "./components/App";
import { FluentProvider, webLightTheme } from "@fluentui/react-components";

/* global document, Office, module, require */

const title = "Contoso Task Pane Add-in";

const rootElement = document.getElementById("container");
const root = createRoot(rootElement);

/* Render application after Office initializes */
Office.onReady(() => {
  root.render(
    <FluentProvider theme={webLightTheme}>
      <App title={title} />
    </FluentProvider>
  );
});

if (module.hot) {
  module.hot.accept("./components/App", () => {
    const NextApp = require("./components/App").default;
    root.render(NextApp);
  });
}

/* React Components */

// SheetCard Component
function SheetCard({ sheetNames }) {
  return (
    <div>
      <h2>Sheet Names</h2>
      <ul>
        {sheetNames.map((name, index) => (
          <li key={index}>{name}</li>
        ))}
      </ul>
    </div>
  );
}

/* Office Add-in Code */

// Handle click event for refresh button
document.getElementById("refresh-button").onclick = () => tryCatch(sheetLoading);

// Load sheet names from Excel and update UI
async function sheetLoading() {
  try {
    await Excel.run(async (context) => {
      const allSheet = context.workbook.worksheets;
      allSheet.load("name");
      await context.sync();

      let sheetNames = [];

      for (let i = 0; i < allSheet.items.length; i++) {
        let current_sheet = allSheet.items[i];
        const sheetName = current_sheet.name;
        sheetNames.push(sheetName);
      }

      console.log(sheetNames);

      renderSheetCard(sheetNames);
    });
  } catch (error) {
    console.error('Error loading sheet names:', error);
  }
}

// Render the SheetCard component
function renderSheetCard(sheetNames) {
  root.render(
    <FluentProvider theme={webLightTheme}>
      <SheetCard sheetNames={sheetNames} />
    </FluentProvider>
  );
}

/** Helper function to catch errors */
async function tryCatch(callback) {
  try {
    await callback();
  } catch (error) {
    console.error(error);
  }
}
