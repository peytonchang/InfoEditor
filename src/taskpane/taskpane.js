/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    // Ensures the DOM is fully loaded before attaching event handlers
    var button = document.getElementById("writeButton");
    button.addEventListener("click", writeValue);
  }
});

function writeValue() {
  // Get the selected field and the input value
  const field = document.getElementById("fieldSelect").value;
  const inputValue = document.getElementById("inputValue").value;

  // Determine the target cell based on the selected field
  let cellAddress;
  switch (field) {
    case "url":
      cellAddress = "B1";
      break;
    case "username":
      cellAddress = "C1";
      break;
    case "ruleService":
      cellAddress = "A1";
      break;
    case "password":
      cellAddress = "D1";
      break;
    default:
      console.error("Invalid selection");
      return;
  }

  // Write the value to the specified cell in the hidden sheet 'info'
  Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem("info");
    const range = sheet.getRange(cellAddress);
    range.values = [[inputValue]]; // Set the value in the selected cell
    await context.sync(); // Synchronize the state with Excel
    console.log(`Value written to ${cellAddress} in 'info' sheet: ${inputValue}`);
  }).catch((error) => {
    console.error("Error:", error);
  });

  // Optional: Clear the input field after writing
  document.getElementById("inputValue").value = "";
}
