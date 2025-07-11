function submitForm(data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Data_Base");  //put the sheet name of yours here 
  const lastRow = sheet.getLastRow();
  const dataRange = sheet.getRange(1, 1, lastRow, sheet.getLastColumn()).getValues();

  let emptyRowIndex = -1;

  for (let i = 0; i < dataRange.length; i++) {
    if (dataRange[i].every(cell => cell === "")) {
      emptyRowIndex = i + 1;
      break;
    }
  }

  if (emptyRowIndex === -1) {
    emptyRowIndex = lastRow + 1;
  }

  
//above this line you don't have to edit any code you have to only focus on below code

  sheet.getRange(emptyRowIndex, 1, 1, 25).setValues([[
    new Date(),
    data.vendorName,
    data.mobileNumber,
    data.suppliedMaterial,
    data.quantity,
    data.uNit
  ]]);
}


//YOU HAVE TO ADD BELOW CODE ALSO SO THAT showForm can it whenever it gets open.

function showForm() {
  const html = HtmlService.createHtmlOutputFromFile("SalesForm")
    .setWidth(400)
    .setHeight(650);
  SpreadsheetApp.getUi().showModalDialog(html, "Add Purchase Entry");
}
// GetUi to add this in to Google sheet Menu

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Open Form")
    .addItem("Entry Form", "showForm")
    //.addItem("Open It", "openGitHubProfile") **You can add more menu by using .addItem in this section**
    .addToUi();
}
