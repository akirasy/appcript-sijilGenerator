// Create topbar menu
function onOpen() {
  SpreadsheetApp.getUi()
  .createMenu('Sijil Generator')
  .addSubMenu(SpreadsheetApp.getUi().createMenu('Setup')
    .addItem('Set Google permission', 'aquireGooglePermission')
    .addItem('Setup Spreadsheet', 'setupSpreadsheet'))
  .addSeparator()
  .addItem('Generate Gdoc File', 'actionGenerateGdoc')
  .addItem('Convert Gdoc to PDF File', 'actionGeneratePdf')
  .addSeparator()
  .addSubMenu(SpreadsheetApp.getUi().createMenu('🕊 About')
      .addItem('⚪ Google AppScript', 'aboutGoogleAppScript')
      .addItem('⚪ Author', 'aboutAuthor')
      .addItem('⚪ License', 'aboutLicense'))
  .addToUi();
}

function aquireGooglePermission() {
  SpreadsheetApp.getUi().alert(
    'Success',
    'If you can see this. You already have permission to use this app.',
    SpreadsheetApp.getUi().ButtonSet.OK);
}

function getConfigVariable() {
  let configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("config.gs");
  let dataSheet   = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Data");
  let workingFolderID       = configSheet.getRange("B6").getValue();
  let templateDocID         = configSheet.getRange("B2").getValue();
  let generatedFolderGdocID = configSheet.getRange("B7").getValue();
  let generatedFolderPdfID  = configSheet.getRange("B8").getValue();

  let output = {
    "dataSheet"           : dataSheet,
    "configSheet"         : configSheet,
    "workingFolder"       : DriveApp.getFolderById(workingFolderID),
    "templateDoc"         : DriveApp.getFileById(templateDocID),
    "generatedFolderGdoc" : DriveApp.getFolderById(generatedFolderGdocID),
    "generatedFolderPdf"  : DriveApp.getFolderById(generatedFolderPdfID),
  }
  return output
}

function setupSpreadsheet() {
  let ui = SpreadsheetApp.getUi();
  let result = ui.alert(
     'Attention!',
     'This should only run once. This function will set up your empty spreadsheet.\n\nAre you sure you want to continue?',
      ui.ButtonSet.YES_NO);
  // Process the user's response.
  let user_response = new Boolean();
  if (result == ui.Button.YES) {
    // User clicked "Yes".
    user_response = true;
  } else {
    // User clicked "No" or X in the title bar.
    ui.alert('Thank you', 'Script exited safely.', ui.ButtonSet.OK);
    user_response = false;
  }

  if (user_response) {
    let sheet1 = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    if (sheet1.getName() == "Sheet1") {
      let activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      let dataSheet   = activeSpreadsheet.insertSheet("Data");
      let configSheet = activeSpreadsheet.insertSheet("config.gs");

      // Configure sheet `Data`
      let a1e1A = dataSheet.getRange("A1:E1");
      a1e1A.merge(); a1e1A.setFontWeight("bold"); a1e1A.setHorizontalAlignment("center"); a1e1A.setBackground("#a4c2f4");
      let f1i1A = dataSheet.getRange("F1:I1");
      f1i1A.merge(); f1i1A.setFontWeight("bold"); f1i1A.setHorizontalAlignment("center"); f1i1A.setBackground("#a4c2f4");
      let a2i2A = dataSheet.getRange("A2:I2");
      a2i2A.setFontWeight("bold"); a2i2A.setHorizontalAlignment("center"); a2i2A.setBackground("#c9daf8");
      let a1i2A = dataSheet.getRange("A1:I2");
      a1i2A.setBorder(true, true, true, true, true, true);
      a1i2A.protect().setWarningOnly(true);
      dataSheet.deleteColumns(10, dataSheet.getMaxColumns() - 9);
      dataSheet.setColumnWidth(1, 300);
      dataSheet.setColumnWidths(2, 4, 150);

      // Configure sheet `config.gs`
      let a1b1B = configSheet.getRange("A1:B1");
      a1b1B.merge(); a1b1B.setFontWeight("bold"); a1b1B.setHorizontalAlignment("center"); a1b1B.setBackground("#c9daf8");
      a1b1B.setBorder(true, true, true, true, true, true);
      a1b1B.protect().setWarningOnly(true);
      let a5b5B = configSheet.getRange("A5:B5");
      a5b5B.merge(); a5b5B.setFontWeight("bold"); a5b5B.setHorizontalAlignment("center"); a5b5B.setBackground("#c9daf8");
      a5b5B.setBorder(true, true, true, true, true, true);
      a5b5B.protect().setWarningOnly(true);
      let a2a3B = configSheet.getRange("A2:A3");
      a2a3B.setFontWeight("bold");
      a2a3B.protect().setWarningOnly(true);
      let a6a8B = configSheet.getRange("A6:A8");
      a6a8B.setFontWeight("bold");
      a6a8B.protect().setWarningOnly(true);
      let b6b8B = configSheet.getRange("B6:B8");
      b6b8B.protect().setWarningOnly(true);
      configSheet.setColumnWidth(1, 200);
      configSheet.setColumnWidth(2, 400);
      configSheet.deleteColumns(3, configSheet.getMaxColumns() - 2);
      configSheet.deleteRows(10, configSheet.getMaxRows() - 9);

      // Prepare directory
      let workingFolder = DriveApp.getFileById(activeSpreadsheet.getId()).getParents().next();
      let generatedFolderGdoc = DriveApp.createFolder("generated-gdoc");
      let generatedFolderPdf = DriveApp.createFolder("generated-pdf");
      generatedFolderGdoc.moveTo(workingFolder);
      generatedFolderPdf.moveTo(workingFolder);

      // Write header values
      dataSheet.getRange("A1").setValue("VARIABLE TAGGING");
      dataSheet.getRange("F1").setValue("OUTPUT LINK");
      dataSheet.getRange("A2").setValue("<<VAR1>>");
      dataSheet.getRange("B2").setValue("<<VAR2>>");
      dataSheet.getRange("C2").setValue("<<VAR3>>");
      dataSheet.getRange("D2").setValue("<<VAR4>>");
      dataSheet.getRange("E2").setValue("<<VAR5>>");
      dataSheet.getRange("F2").setValue("GDOC URL");
      dataSheet.getRange("G2").setValue("GDOC ID");
      dataSheet.getRange("H2").setValue("PDF URL");
      dataSheet.getRange("I2").setValue("PDF ID");

      configSheet.getRange("A1").setValue("USER-DEFINED VALUES");
      configSheet.getRange("A2").setValue("Template file ID");
      configSheet.getRange("A5").setValue("AUTO-GENERATED VALUES");
      configSheet.getRange("A6").setValue("Working folder ID");
      configSheet.getRange("A7").setValue("GoogleDoc folder ID");
      configSheet.getRange("A8").setValue("PDF folder ID");
      configSheet.getRange("B6").setValue(workingFolder.getId());
      configSheet.getRange("B7").setValue(generatedFolderGdoc.getId());
      configSheet.getRange("B8").setValue(generatedFolderPdf.getId());

      // Delete sheet1
      activeSpreadsheet.deleteSheet(sheet1);

      // Prompt complete instruction
      ui.alert("Success!", "Please enter your Sijil template in config.gs sheet.", ui.ButtonSet.OK);
    } else {
      ui.alert("Attention!", "This is not an empty spreadsheet.\nPlease create new spreadsheet.\n\nScript will abort.", ui.ButtonSet.OK);
    }
  }
}

function generateGdocFile(generatedFolderGdoc, templateDoc, varTag) {
  let newDocObj = templateDoc.makeCopy(generatedFolderGdoc);
  let body = DocumentApp.openById(newDocObj.getId()).getBody();
  body.replaceText("<<VAR1>>", varTag[0]);
  body.replaceText("<<VAR2>>", varTag[1]);
  body.replaceText("<<VAR3>>", varTag[2]);
  body.replaceText("<<VAR4>>", varTag[3]);
  body.replaceText("<<VAR5>>", varTag[4]);
  newDocObj.setName(varTag[0]);
  return newDocObj
}

function convertToPdf(gdocFile, targetFolderID) {
  let docBlob = gdocFile.getAs('application/pdf');
  docBlob.setName(gdocFile.getName() + ".pdf");
  let convertedPdf = DriveApp.createFile(docBlob).moveTo(targetFolderID);
  return convertedPdf
}

function actionGenerateGdoc() {
  let confVar = getConfigVariable();

  let generatedFolderGdoc = confVar.generatedFolderGdoc;
  let templateDoc = confVar.templateDoc;
  let sheet = confVar.dataSheet;

  // generate and collect info
  let candidateArray = new Array(); // [index, url, fileID]
  let workingRangeValues = sheet.getRange(3, 1, sheet.getLastRow() - 2, 8).getValues();
  workingRangeValues.map((item, index) => {
    if (item[5] == "") {
      let candidateFile = generateGdocFile(generatedFolderGdoc, templateDoc, item);
      let candidateIndex = index + 3;
      candidateArray.push([candidateIndex, candidateFile.getUrl(), candidateFile.getId()]);
      Logger.log("File created: " + item[0]);
    }
  })

  // write url to cell
  candidateArray.forEach(item => {
    sheet.getRange(item[0], 6).setValue(item[1]);
    sheet.getRange(item[0], 7).setValue(item[2]);
  })
}

function actionGeneratePdf() {
  let confVar = getConfigVariable();
  let targetFolder = confVar.generatedFolderPdf;
  let sheet = confVar.dataSheet;

  let candidateArray = new Array(); // [index, url, fileID]
  let workingRangeValues = sheet.getRange(3, 1, sheet.getLastRow() - 2, 8).getValues();
  workingRangeValues.map((item, index) => {
    if (item[7] == "") {
      let gdocFile = DriveApp.getFileById(item[6]);
      let candidateIndex = index + 3;
      let convertedPdf = convertToPdf(gdocFile, targetFolder);
      candidateArray.push([candidateIndex, convertedPdf.getUrl(), convertedPdf.getId()])
      Logger.log("File converted: " + gdocFile.getName());
    }
  })

  // write url to cell
  candidateArray.forEach(item => {
    sheet.getRange(item[0], 8).setValue(item[1]);
    sheet.getRange(item[0], 9).setValue(item[2]);
  })
}

function aboutLicense() {
  let title = 'Open Source';
  let subtitle = `
    This app is open source and free to use under the terms of GNU General Public License v3.0.

    Copyright (C) 2021  akirasy
    This program is free software: you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation, either version 3 of the License, or
    (at your option) any later version.
    This program is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU General Public License for more details.
    You should have received a copy of the GNU General Public License
    along with this program.  If not, see <https://www.gnu.org/licenses/>.
  `;
  SpreadsheetApp.getUi().alert(title, subtitle, SpreadsheetApp.getUi().ButtonSet.OK);
}

function aboutAuthor() {
  let title = 'AppScript Author';
  let subtitle = `
    This app is developed by akirasy <fitri.abakar@gmail.com>
    
    Feel free to browse other app here --> https://github.com/akirasy
    For this specific app source, look here --> https://github.com/akirasy/appcript-sijilGenerator
  `;
  SpreadsheetApp.getUi().alert(title, subtitle, SpreadsheetApp.getUi().ButtonSet.OK);
}

function aboutGoogleAppScript() {
  let title = 'Google AppScript';
  let subtitle = `
    Google Apps Script is a rapid application development platform that makes it 
    fast and easy to create business applications that integrate with Google Workspace. 
    
    You write code in modern JavaScript and have access to built-in libraries for favorite 
    Google Workspace applications like Gmail, Calendar, Drive, and more. 
    
    There's nothing to install—we give you a code editor right in your browser, 
    and your scripts run on Google's servers.

    Learn more at --> https://developers.google.com/apps-script/overview
  `;
  SpreadsheetApp.getUi().alert(title, subtitle, SpreadsheetApp.getUi().ButtonSet.OK);
}
