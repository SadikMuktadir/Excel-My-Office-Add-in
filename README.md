
# Navigator App

This website is simple My office add-in project , where you can go up and down the sheet very easily.


## Installation

Install my-project with npm

```bash
  npm install -g yo generator-office
  yo office

```
    
## Deployment

To deploy this project run

```bash
  cd "My Office Add-in"
  npm run dev-server
  npm start
```


## Usage/Examples

```javascript
import * as React from "react";
import { createRoot } from "react-dom/client";
import App from "./components/App";
import { FluentProvider, webLightTheme } from "@fluentui/react-components";
import { LiaArrowUpSolid, LiaArrowDownSolid } from "react-icons/lia";
import { TbArrowsDown, TbArrowsUp } from "react-icons/tb";
import { Button } from "@fluentui/react-components";
import { ButtonProps } from "@fluentui/react-components";


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

                <Button
                  style={{ marginRight: "5px" }}
                  shape="square"
                  Appearance="primary"
                  size="small"
                  onClick={() => moveSheetUp(index)}
                  disabled={index === 0} 
                >
                  <LiaArrowUpSolid />
                </Button>

                <Button
                  style={{ marginRight: "5px" }}
                  shape="square"
                  Appearance="primary"
                  size="small"
                  onClick={() => moveSheetDown(index)}
                  disabled={index === sheetNames.length - 1}
                >
                  <LiaArrowDownSolid />
                </Button>
                <Button
                  style={{ marginRight: "5px" }}
                  shape="square"
                  Appearance="primary"
                  size="small"
                  onClick={() => moveAllSheetsUp(index)}
                  disabled={index === 0}
                >
                  <TbArrowsUp />
                </Button>

                <Button
                  style={{ marginRight: "5px" }}
                  shape="square"
                  Appearance="primary"
                  size="small"
                  onClick={() => moveAllSheetsDown(index)}
                  disabled={index === sheetNames.length - 1} 
                >
                  <TbArrowsDown />
                </Button>
                <div>{sheetName}</div>
              </div>
            </li>
          ))}
        </ul>

        <div style={{ marginLeft: "40px" }}>
          <Button
            onClick={() => onReorderButtonClick(sheetNames)}
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
          </Button>
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

async function sheetLoading() {
  try {
    await Excel.run(async (context) => {
      const allSheet = context.workbook.worksheets;
      allSheet.load(["name", "visibility"]); 
      await context.sync();

      let sheetNames = [];

      for (let i = 0; i < allSheet.items.length; i++) {
        let current_sheet = allSheet.items[i];
        const sheetName = current_sheet.name;
        const isHidden = current_sheet.visibility === "Hidden";
        if (!isHidden) {
         
          sheetNames.push(sheetName);
        }
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


      worksheets.load(["name", "visibility"]);
      await context.sync();

    
      const visibleSheetNames = sheetNames.filter(sheetName => {
        const worksheet = worksheets.items.find(sheet => sheet.name === sheetName);
        return worksheet.visibility !== "Hidden";
      });


      const worksheetMap = {};
      for (let i = 0; i < worksheets.items.length; i++) {
        const worksheet = worksheets.items[i];
        worksheetMap[worksheet.name] = worksheet;
      }

     
      for (let i = 0; i < visibleSheetNames.length; i++) {
        const sheetName = visibleSheetNames[i];
        const worksheet = worksheetMap[sheetName];
        worksheet.position = i; 
      }


      await context.sync();

   
      const updatedSheetNames = [];
      for (let i = 0; i < worksheets.items.length; i++) {
        const currentSheet = worksheets.items[i];
        if (currentSheet.visibility !== "Hidden") {
          updatedSheetNames.push(currentSheet.name);
        }
      }

      renderSheetCard(updatedSheetNames);
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

}
```


## Authors

- [@SadikMuktadir](https://www.github.com/SadikMuktadir)
- [@Ismail Hosen](https://www.github.com/1504168)


## Documentation

[Excel task pane add-in](https://learn.microsoft.com/en-us/office/dev/add-ins/quickstarts/excel-quickstart-react?fbclid=IwAR31M4sUWf29qaeji4NeqqxqTDb_EuS9J3DL_HmjK8jVWBleQH1ZtULJXEs)

[React JS](https://react.dev/)

[Fluent UI](https://react.fluentui.dev/?path=/docs/components-datagrid--default)

![image](https://github.com/SadikMuktadir/Excel-My-Office-Add-in/assets/136139497/ba77f1cb-03e3-4c69-beaa-696a56df9bc4)

