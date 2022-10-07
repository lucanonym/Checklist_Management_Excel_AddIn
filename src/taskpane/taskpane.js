/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    // Determine if the user's version of Office supports all the Office.js APIs that are used in the tutorial.
if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
  console.log('Sorry. This add-in uses Excel.js APIs that are not available in your version of Office.');
}

// Assign event handlers and other initialization logic.
document.getElementById("delete-checkboxes").onclick = testEineTabelle;
document.getElementById("sideload-msg").style.display = "none";
document.getElementById("app-body").style.display = "flex";

  }
});
async function testEineTabelle() {
  await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getActiveWorksheet();
    let finishedSheet = sheet.copy(Excel.WorksheetPositionType.after, sheet);
    finishedSheet.name = 'FinishedCalculation';
    await context.sync();


    // hiding config and checkboxes

    finishedSheet.getRange('F:F').columnHidden = true;
    finishedSheet.getRange('I:I').columnHidden = true;
    finishedSheet.getRange('H:H').columnHidden = true;

    

    // filtering rows

    let table = finishedSheet.tables.load('tables');
    let filteringTable = table.getItemAt(0);
    filterTable(filteringTable);
    await context.sync();
    
  })
  .catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
  });


function filterTable(firstTable) {
  let amountFilter = firstTable.columns.getItem('Config').filter;
  amountFilter.apply({
    filterOn: Excel.FilterOn.custom,
    criterion1: ">=1",
  });
}
}

