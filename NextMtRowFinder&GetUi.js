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
