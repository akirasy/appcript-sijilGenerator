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
  .addItem('Send email to selected candidate', 'actionSendEmail')
  .addSubMenu(SpreadsheetApp.getUi().createMenu('🕊 About')
      .addItem('⚪ Google AppScript', 'aboutGoogleAppScript')
      .addItem('⚪ Author', 'aboutAuthor')
      .addItem('⚪ License', 'aboutLicense'))
  .addToUi();
}

/**
 * Check if user has approve script execution.
 */
function aquireGooglePermission() {
  SpreadsheetApp.getUi().alert(
    'Success',
    'If you can see this. You already have permission to use this app.',
    SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Script to set up empty spreadsheet.
 */
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
      Logger.log("Preparing new sheet: Data");
      let a1e1A = dataSheet.getRange("A1:E1");
      a1e1A.merge(); a1e1A.setFontWeight("bold"); a1e1A.setHorizontalAlignment("center"); a1e1A.setBackground("#a4c2f4");
      let f1i1A = dataSheet.getRange("F1:I1");
      f1i1A.merge(); f1i1A.setFontWeight("bold"); f1i1A.setHorizontalAlignment("center"); f1i1A.setBackground("#a4c2f4");
      let j1k1A = dataSheet.getRange("J1:K1");
      j1k1A.merge(); j1k1A.setFontWeight("bold"); j1k1A.setHorizontalAlignment("center"); j1k1A.setBackground("#a4c2f4");
      let a2k2A = dataSheet.getRange("A2:K2");
      a2k2A.setFontWeight("bold"); a2k2A.setHorizontalAlignment("center"); a2k2A.setBackground("#c9daf8");
      let a1k2A = dataSheet.getRange("A1:K2");
      a1k2A.setBorder(true, true, true, true, true, true);
      a1k2A.protect().setWarningOnly(true);
      dataSheet.deleteColumns(12, dataSheet.getMaxColumns() - 11);
      dataSheet.setColumnWidth(1, 300);
      dataSheet.setColumnWidths(2, 4, 150);

      // Configure sheet `config.gs`
      Logger.log("Preparing new sheet: config.gs");
      let a1b1B = configSheet.getRange("A1:B1");
      a1b1B.merge(); a1b1B.setFontWeight("bold"); a1b1B.setHorizontalAlignment("center"); a1b1B.setBackground("#c9daf8");
      a1b1B.setBorder(true, true, true, true, true, true);
      a1b1B.protect().setWarningOnly(true);
      let a6b6B = configSheet.getRange("A6:B6");
      a6b6B.merge(); a6b6B.setFontWeight("bold"); a6b6B.setHorizontalAlignment("center"); a6b6B.setBackground("#c9daf8");
      a6b6B.setBorder(true, true, true, true, true, true);
      a6b6B.protect().setWarningOnly(true);
      let a2a4B = configSheet.getRange("A2:A4");
      a2a4B.setFontWeight("bold");
      a2a4B.protect().setWarningOnly(true);
      let a7a9B = configSheet.getRange("A7:A9");
      a7a9B.setFontWeight("bold");
      a7a9B.protect().setWarningOnly(true);
      let b7b9B = configSheet.getRange("B7:B9");
      b7b9B.protect().setWarningOnly(true);
      configSheet.setColumnWidth(1, 200);
      configSheet.setColumnWidth(2, 400);
      configSheet.deleteColumns(3, configSheet.getMaxColumns() - 2);
      configSheet.deleteRows(10, configSheet.getMaxRows() - 9);

      // Prepare directory
      Logger.log("Setting up file directory in parent folder");
      let workingFolder = DriveApp.getFileById(activeSpreadsheet.getId()).getParents().next();
      let generatedFolderGdoc = DriveApp.createFolder("generated-gdoc");
      let generatedFolderPdf = DriveApp.createFolder("generated-pdf");
      generatedFolderGdoc.moveTo(workingFolder);
      generatedFolderPdf.moveTo(workingFolder);

      // Write header values
      Logger.log("Editing and formatting sheet: Data");
      dataSheet.getRange("A1").setValue("VARIABLE TAGGING");
      dataSheet.getRange("F1").setValue("OUTPUT LINK");
      dataSheet.getRange("J1").setValue("EMAIL RECIPIENT");
      dataSheet.getRange("A2").setValue("<<VAR1>>");
      dataSheet.getRange("B2").setValue("<<VAR2>>");
      dataSheet.getRange("C2").setValue("<<VAR3>>");
      dataSheet.getRange("D2").setValue("<<VAR4>>");
      dataSheet.getRange("E2").setValue("<<VAR5>>");
      dataSheet.getRange("F2").setValue("GDOC URL");
      dataSheet.getRange("G2").setValue("GDOC ID");
      dataSheet.getRange("H2").setValue("PDF URL");
      dataSheet.getRange("I2").setValue("PDF ID");
      dataSheet.getRange("J2").setValue("EMAIL");
      dataSheet.getRange("K2").setValue("SENT");

      Logger.log("Editing and formatting sheet: config.gs");
      configSheet.getRange("A1").setValue("USER-DEFINED VALUES");
      configSheet.getRange("A2").setValue("Template file ID");
      configSheet.getRange("A3").setValue("Subject for email");
      configSheet.getRange("A4").setValue("Message for email");
      
      configSheet.getRange("A6").setValue("AUTO-GENERATED VALUES");
      configSheet.getRange("A7").setValue("Working folder ID");
      configSheet.getRange("A8").setValue("GoogleDoc folder ID");
      configSheet.getRange("A9").setValue("PDF folder ID");
      configSheet.getRange("B7").setValue(workingFolder.getId());
      configSheet.getRange("B8").setValue(generatedFolderGdoc.getId());
      configSheet.getRange("B9").setValue(generatedFolderPdf.getId());

      // Delete sheet1
      Logger.log("Deleting default sheet1");
      activeSpreadsheet.deleteSheet(sheet1);

      // Prompt complete instruction
      Logger.log("Prompt user to fill in GDoc Template ID");
      ui.alert("Success!", "Please enter your Sijil template in config.gs sheet.", ui.ButtonSet.OK);
    } else {
      ui.alert("Attention!", "This is not an empty spreadsheet.\nPlease create new spreadsheet.\n\nScript will abort.", ui.ButtonSet.OK);
    }
  }
}

/**
 * Script for onOpen() - generate sijil on selected row
 */
function actionGenerateGdoc() {
  let confVar             = getConfigVariable();
  let generatedFolderGdoc = confVar.generatedFolderGdoc;
  let templateDoc         = confVar.templateDoc;
  let sheet               = confVar.dataSheet;
  let selectedRange       = sheet.getActiveRange();

  let forloopStart = selectedRange.getRowIndex();
  let forloopEnd   = forloopStart + selectedRange.getNumRows();
  for (let rowid = forloopStart; rowid < forloopEnd; rowid++) {
    let candidateVarTag = sheet.getRange(rowid, 1, 1, 5).getValues()[0];
    let candidateFile = generateGdocFile(generatedFolderGdoc, templateDoc, candidateVarTag);
    sheet.getRange(rowid, 6).setValue(candidateFile.getUrl());
    sheet.getRange(rowid, 7).setValue(candidateFile.getId());
    Logger.log("Rowid: " + rowid + " - File created: " + candidateVarTag[0]);
  }
}

/**
 * Script for onOpen() - convert sijil from Gdoc to PDF on selected row
 */
function actionGeneratePdf() {
  let confVar       = getConfigVariable();
  let sheet         = confVar.dataSheet;
  let targetFolder  = confVar.generatedFolderPdf;
  let selectedRange = sheet.getActiveRange();

  let forloopStart = selectedRange.getRowIndex();
  let forloopEnd   = forloopStart + selectedRange.getNumRows();
  for (let rowid = forloopStart; rowid < forloopEnd; rowid++) {
    let gdocId       = sheet.getRange(rowid, 7).getValue();
    let gdocFile     = DriveApp.getFileById(gdocId);
    let convertedPdf = convertToPdf(gdocFile, targetFolder);
    sheet.getRange(rowid, 8).setValue(convertedPdf.getUrl());
    sheet.getRange(rowid, 9).setValue(convertedPdf.getId());
    Logger.log("File converted for rowid: " + rowid);
  }
}

/**
 * Script for onOpen() - send email with PDF attachment on selected row
 */
function actionSendEmail() {
  let confVar       = getConfigVariable();
  let sheet         = confVar.dataSheet;
  let selectedRange = sheet.getActiveRange();

  let forloopStart = selectedRange.getRowIndex();
  let forloopEnd = forloopStart + selectedRange.getNumRows();
  for (let rowid = forloopStart; rowid < forloopEnd; rowid++) {
    let recipientEmail = sheet.getRange(rowid, 10).getValue();
    let recipientAttachmentId = sheet.getRange(rowid, 9).getValue();
    sendEmailTo(confVar, recipientEmail, recipientAttachmentId);
    sheet.getRange(rowid, 11).setValue('DONE');
    Logger.log('Email sent to: ' + recipientEmail);
  };
}

/**
 * Instantiate global project variable to save execution time.
 */
function getConfigVariable() {
  let configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("config.gs");
  let dataSheet   = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Data");
  let workingFolderID       = configSheet.getRange("B7").getValue();
  let templateDocID         = configSheet.getRange("B2").getValue();
  let generatedFolderGdocID = configSheet.getRange("B8").getValue();
  let generatedFolderPdfID  = configSheet.getRange("B9").getValue();
  let emailSubject          = configSheet.getRange("B3").getValue();
  let emailBody             = configSheet.getRange("B4").getValue();

  let output = {
    "dataSheet"           : dataSheet,
    "configSheet"         : configSheet,
    "workingFolder"       : DriveApp.getFolderById(workingFolderID),
    "templateDoc"         : DriveApp.getFileById(templateDocID),
    "generatedFolderGdoc" : DriveApp.getFolderById(generatedFolderGdocID),
    "generatedFolderPdf"  : DriveApp.getFolderById(generatedFolderPdfID),
    'emailSubject'        : emailSubject,
    'emailBody'           : emailBody
  }
  return output
}

/**
 * Generate a single sijil as GoogleDoc object.
 * @param {Object} generatedFolderGdoc target folder for generated file
 * @param {Object} templateDoc GoogleDoc template file to copy from 
 * @param {Array} varTag list of variableTag from activeSpreadsheet
 */
function generateGdocFile(generatedFolderGdoc, templateDoc, varTag) {
  let newDocObj = templateDoc.makeCopy(generatedFolderGdoc);
  let docApp = DocumentApp.openById(newDocObj.getId());
  let body = docApp.getBody();
  body.replaceText("<<VAR1>>", varTag[0]);
  body.replaceText("<<VAR2>>", varTag[1]);
  body.replaceText("<<VAR3>>", varTag[2]);
  body.replaceText("<<VAR4>>", varTag[3]);
  body.replaceText("<<VAR5>>", varTag[4]);
  if ( body.findText("<<HASHID>>") ) {
    body.replaceText("<<HASHID>>", newDocObj.getId());
  }
  newDocObj.setName(varTag[0]);
  return newDocObj
}

/**
 * Convert GoogleDoc object file into PDF.
 * @param {Object} gdocFile GoogleDoc file to convert
 * @param {Object} targetFolder target folder for generated file
 */
function convertToPdf(gdocFile, targetFolder) {
  let docBlob = gdocFile.getAs('application/pdf');
  docBlob.setName(gdocFile.getName() + ".pdf");
  let convertedPdf = DriveApp.createFile(docBlob).moveTo(targetFolder);
  return convertedPdf
}

/**
 * Send email with sijil attachment.
 * @param {Object} confVar instance of getConfigVariable()
 * @param {string} recipientEmail
 * @param {string} fileAttachmentId
 */
function sendEmailTo(confVar, recipientEmail, fileAttachmentId) {
  let fileBlob     = DriveApp.getFileById(fileAttachmentId).getBlob();
  MailApp.sendEmail({
    to          : recipientEmail,
    subject     : confVar.emailSubject,
    body        : confVar.emailBody,
    attachments : [fileBlob]
  });
}

/**
 * Prompt user about license.
 */
function aboutLicense() {
  let title = 'Open Source';
  let subtitle = `
    This app is open source and free to use under the terms of GNU General Public License v3.0.

    Copyright (C) 2023  akirasy
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

/**
 * Prompt user about author.
 */
function aboutAuthor() {
  let title = 'AppScript Author';
  let subtitle = `
    This app is developed by akirasy <fitri.abakar@gmail.com>
    
    Feel free to browse other app here --> https://github.com/akirasy
    For this specific app source, look here --> https://github.com/akirasy/appcript-sijilGenerator
  `;
  SpreadsheetApp.getUi().alert(title, subtitle, SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Prompt user about GoogleAppScript.
 */
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
