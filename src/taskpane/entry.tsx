import * as React from "react";
import * as ReactDOM from "react-dom/client";
import TaskPane from "./TaskPane";
import { FluentProvider, webLightTheme } from "@fluentui/react-components";

// Safe Office loader
function renderTaskpane() {
  const container = document.getElementById("container");
  if (container) {
    const root = ReactDOM.createRoot(container);
    root.render(
      <FluentProvider theme={webLightTheme}>
        <TaskPane />
      </FluentProvider>
    );
  }
}

// Ensure Office.js is available before running
function waitForOfficeReady() {
  if (typeof Office !== "undefined" && Office.onReady) {
    Office.onReady(() => {
      console.log("✅ Office.js ready – rendering taskpane");
      renderTaskpane();
    });
  } else {
    console.warn("⚠️ Office not yet defined, retrying...");
    setTimeout(waitForOfficeReady, 300);
  }
}

waitForOfficeReady();
