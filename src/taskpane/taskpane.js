/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    // Ensures the DOM is fully loaded before attaching event handlers
    var write = document.getElementById("writeButton");
    var read = document.getElementById("readButton");

    write.addEventListener("click", writeValue);
    read.addEventListener("click", readValue);
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

function readValue() {
  const field = document.getElementById("fieldSelect").value;
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

  Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem("info");
    const sourceRange = sheet.getRange(cellAddress);
    sourceRange.load("values");

    await context.sync();

    // Now that we have the value, write it to A1 of the active sheet
    const currentSheet = context.workbook.worksheets.getActiveWorksheet();
    const targetRange = currentSheet.getRange("A1");
    targetRange.values = sourceRange.values;

    await context.sync();
    console.log(`Value from ${cellAddress} read and written to A1 of the active sheet.`);
  }).catch((error) => {
    console.error("Error:", error);
  });
}
