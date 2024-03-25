import * as React from "react";
import { createRoot } from "react-dom/client";
import App from "./components/App";
import { FluentProvider, webLightTheme } from "@fluentui/react-components";
import { LiaArrowUpSolid, LiaArrowDownSolid } from "react-icons/lia";
import { TbArrowsDown, TbArrowsUp } from "react-icons/tb";

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
// Functionality

/* React Components */

// SheetCard Component
function SheetCard({ sheetNames, updateSheetOrder, onReorderButtonClick }) {
  const moveSheetUp = (index) => {
    if (index > 0) {
      const updatedOrder = [...sheetNames];
      const temp = updatedOrder[index - 1];
      updatedOrder[index - 1] = updatedOrder[index];
      updatedOrder[index] = temp;
      updateSheetOrder(updatedOrder);
    }
  };

  const moveSheetDown = (index) => {
    if (index < sheetNames.length - 1) {
      const updatedOrder = [...sheetNames];
      const temp = updatedOrder[index + 1];
      updatedOrder[index + 1] = updatedOrder[index];
      updatedOrder[index] = temp;
      updateSheetOrder(updatedOrder);
    }
  };

  const moveAllSheetsUp = (index) => {
    const updatedOrder = [...sheetNames];
    const sheet = updatedOrder.splice(index, 1);
    updatedOrder.unshift(sheet[0]);
    updateSheetOrder(updatedOrder);
  };

  const moveAllSheetsDown = (index) => {
    const updatedOrder = [...sheetNames];
    const sheet = updatedOrder.splice(index, 1);
    updatedOrder.push(sheet[0]);
    updateSheetOrder(updatedOrder);
  };

  return (
    <div style={{ backgroundColor: "white", height: "500px", padding: "10px" }}>
      <div>
        <h1 style={{ color: "#124076", marginLeft: "40px" }}>Navigator App</h1>
      </div>
      <div>
        <h2 style={{ color: "#4CCD99", marginLeft: "40px" }}>Total Sheets: {sheetNames.length}</h2>
        <ul type="none">
          {sheetNames.map((sheetName, index) => (
            <li key={index}>
              <div style={{ display: "flex", alignItems: "center" }}>
                {/* Move Up Button */}
                <button
                  style={{ marginRight: "5px", cursor: "pointer" }}
                  onClick={() => moveSheetUp(index)}
                  disabled={index === 0} // Disable if it's the first sheet
                >
                  <LiaArrowUpSolid />
                </button>
                {/* Move Down Button */}
                <button
                  style={{ marginRight: "5px", cursor: "pointer" }}
                  onClick={() => moveSheetDown(index)}
                  disabled={index === sheetNames.length - 1} // Disable if it's the last sheet
                >
                  <LiaArrowDownSolid />
                </button>
                <button
                  style={{ marginRight: "5px", cursor: "pointer" }}
                  onClick={() => moveAllSheetsUp(index)}
                  disabled={index === 0}
                >
                  <TbArrowsUp />
                </button>
                {/* Move Down All Button */}
                <button
                  style={{ marginRight: "5px", cursor: "pointer" }}
                  onClick={() => moveAllSheetsDown(index)}
                  disabled={index === sheetNames.length - 1} // Disable if it's the last sheet
                >
                  <TbArrowsDown />
                </button>
                <div>{sheetName}</div>
              </div>
            </li>
          ))}
        </ul>

        <div style={{ marginLeft: "40px" }}>
          <button
           onClick={() => onReorderButtonClick(sheetNames)}
            style={{ padding: "10px", fontWeight: "bold", cursor: "pointer" }}
            onMouseOver={(e) => {
              e.target.style.backgroundColor = "#FFF";
              e.target.style.color = "#40679E";
            }}
            onMouseOut={(e) => {
              e.target.style.backgroundColor = "#40679E";
              e.target.style.color = "#FFF";
            }}
          >
            Re-Order Sheet
          </button>
        </div>
      </div>
    </div>
  );
}

/* Office Add-in Code */

// Handle click event for refresh button
// document.getElementById("refresh-button").onclick = () => tryCatch(sheetLoading);

Office.onReady(() => {
  tryCatch(sheetLoading);
});

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
    console.error("Error loading sheet names:", error);
  }
}

const onReorderButtonClick = async (sheetNames) => {
  try {
    await Excel.run(async (context) => {
      const workbook = context.workbook;
      const worksheets = workbook.worksheets;

      // Get the current order of worksheets in the workbook
      const currentSheetNames = sheetNames.slice(); // Create a copy of sheetNames array
      const currentWorksheets = worksheets.load("name");
      await context.sync();

      // Map the current order of sheet names to their corresponding worksheet objects
      const worksheetMap = {};
      for (let i = 0; i < currentWorksheets.items.length; i++) {
        const worksheet = currentWorksheets.items[i];
        worksheetMap[worksheet.name] = worksheet;
      }

      // Reorder worksheets based on the current UI order
      for (let i = 0; i < currentSheetNames.length; i++) {
        const sheetName = currentSheetNames[i];
        const worksheet = worksheetMap[sheetName];
        worksheet.position = i; // Set the position of the worksheet
      }

      // Save the changes
      await context.sync();
    });
  } catch (error) {
    console.error("Error reordering sheets:", error);
  }
};

const updateSheetOrder = (updatedOrder) => {
  tryCatch(() => {
    renderSheetCard(updatedOrder);
  });
};

function renderSheetCard(sheetNames) {
  root.render(
    <FluentProvider theme={webLightTheme}>
      <SheetCard
        sheetNames={sheetNames}
        updateSheetOrder={updateSheetOrder}
        onReorderButtonClick={onReorderButtonClick}
      />
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
