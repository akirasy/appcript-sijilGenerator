function onOpen() {
  SpreadsheetApp.getUi()
  .createMenu('Sijil Generator')
  .addItem('Uppercase', 'toUpperCase')
  .addSubMenu(SpreadsheetApp.getUi().createMenu('Setup')
    .addItem('Set Google permission', 'aquireUserPermission')
    .addItem('Setup Spreadsheet', 'setupSpreadsheet'))
  .addSeparator()
  .addItem('Generate Gdoc File', 'actionGenerateGdoc')
  .addItem('Convert Gdoc to PDF File', 'actionGeneratePdf')
  .addItem('Send email to selected candidate', 'actionSendEmail')
  .addSeparator()
  .addSubMenu(SpreadsheetApp.getUi().createMenu('🕊 About')
      .addItem('⚪ Google AppScript', 'aboutGoogleAppScript')
      .addItem('⚪ Author', 'aboutAuthor')
      .addItem('⚪ License', 'aboutLicense'))
  .addToUi();
}

/**
 * Check if user has allow permission to run this app.
 */
function aquireUserPermission() {
  SpreadsheetApp.getUi().alert(
    'Success',
    'If you can see this. You already have permission to use this app.',
    SpreadsheetApp.getUi().ButtonSet.OK);
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
    dataSheet           : dataSheet,
    configSheet         : configSheet,
    workingFolder       : DriveApp.getFolderById(workingFolderID),
    templateDoc         : DriveApp.getFileById(templateDocID),
    generatedFolderGdoc : DriveApp.getFolderById(generatedFolderGdocID),
    generatedFolderPdf  : DriveApp.getFolderById(generatedFolderPdfID),
    emailSubject        : emailSubject,
    emailBody           : emailBody
  }
  return output
}

/**
 * Script to set up empty spreadsheet.
 */
function setupSpreadsheet() {
  
  function promptSetupSpreadsheet(ui) {
    let result = ui.alert(
      'Attention!',
      'This should only run once. This function will set up your empty spreadsheet.\n\nAre you sure you want to continue?',
        ui.ButtonSet.YES_NO);
    // Process the user's response.
    let userResponse = new Boolean();
    if (result == ui.Button.YES) {
      // User clicked "Yes".
      userResponse = true;
    } else {
      // User clicked "No" or X in the title bar.
      ui.alert('Thank you', 'Script exited safely.', ui.ButtonSet.OK);
      userResponse = false;
    }
    return userResponse
  }

  function configureDataSheet(sheetObj) {
    Logger.log('Preparing new sheet: Data');
    sheetObj.getRange('A1:Q2').setValues(
      [[
        'GENERATE GDOCS', '',
        'GENERATE PDF', '',
        'EMAIL RECIPIENT', '',
        'BUFFER', '',
        'ADD YOUR CUSTOM VARIABLE HERE. YOU MAY ADD AS MANY COLUMNS (VARIABLES) AS YOU WANT',
        '', '', '', '', '', '', '', ''
       ],
       [
        'GDOC URL', 'GDOC ID'  , 'PDF URL' , 'PDF ID'  ,
        'EMAIL'   , 'SENT'     , 'BUFFER1' , 'BUFFER2' , 
        '<<VAR1>>', '<<VAR2>>' , '<<VAR3>>', '<<VAR4>>', 
        '<<VAR5>>', '<<VAR6>>' , '<<VAR7>>', '<<VAR8>>',
        '<<VAR9>>'
       ]]
    );
    sheetObj.getRange('A1:B1').merge();
    sheetObj.getRange('C1:D1').merge();
    sheetObj.getRange('E1:F1').merge();
    sheetObj.getRange('G1:H1').merge();

    sheetObj.getRange('1:1').setBackground('#a4c2f4');
    sheetObj.getRange('2:2').setBackground('#c9daf8');
    sheetObj.getRange('A3:D').setBackground('#d9d9d9');
    sheetObj.getRange('F3:H').setBackground('#d9d9d9');

    sheetObj.getRange('1:2')
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setBorder(true, true, true, true, true, true);
    sheetObj.getRange('I1').setHorizontalAlignment('left');
  }

  function configureConfigSheet(sheetObj) {
    Logger.log("Preparing new sheet: config.gs");
    let headerValues = [
      ['USER-DEFINED VALUES'],
      ['Template file ID'],
      ['Subject for email'],
      ['Message for email'],
      [''],
      ['AUTO-GENERATED VALUES'],
      ['Working folder ID'],
      ['GoogleDoc folder ID'],
      ['PDF folder ID']
    ]

    sheetObj.setColumnWidth(1, 200);
    sheetObj.setColumnWidth(2, 400);
    sheetObj.deleteColumns(3, sheetObj.getMaxColumns() - 2);
    sheetObj.deleteRows(10, sheetObj.getMaxRows() - 9);

    sheetObj.getRange('A1:A9').setValues(headerValues);
    sheetObj.getRange('A1:A4')
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setBorder(true, true, true, true, true, true)
      .setBackground('#c9daf8')
      .protect().setWarningOnly(true);
    sheetObj.getRange('A6:A9')
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setBorder(true, true, true, true, true, true)
      .setBackground('#c9daf8')
      .protect().setWarningOnly(true);
    sheetObj.getRange('A1:B1').merge().setBackground('#a4c2f4');
    sheetObj.getRange('A6:B6').merge().setBackground('#a4c2f4');
  }

  function prepareDirectory(activeSpreadsheet, configSheet) {
    Logger.log('Setting up file directory in parent folder');
    let workingFolder = DriveApp.getFileById(activeSpreadsheet.getId()).getParents().next();
    let generatedFolderGdoc = DriveApp.createFolder('generated-gdoc');
    let generatedFolderPdf = DriveApp.createFolder('generated-pdf');
    generatedFolderGdoc.moveTo(workingFolder);
    generatedFolderPdf.moveTo(workingFolder);

    configSheet.getRange('B7').setValue(workingFolder.getId());
    configSheet.getRange('B8').setValue(generatedFolderGdoc.getId());
    configSheet.getRange('B9').setValue(generatedFolderPdf.getId());
  }

  function promptFillTemplate(ui) {
    // Prompt complete instruction
    Logger.log("Prompt user to fill in GDoc Template ID");
    ui.alert("Success!", "Please enter your Sijil template in config.gs sheet.", ui.ButtonSet.OK);
  }

  // Run
  let ui = SpreadsheetApp.getUi();
  let userResponse = promptSetupSpreadsheet(ui);
  if (userResponse) {
    let sheet1 = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    if (sheet1.getName() == "Sheet1") {
      let activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      let dataSheet   = sheet1.setName('Data');
      let configSheet = activeSpreadsheet.insertSheet('config.gs');

      configureDataSheet(dataSheet);
      configureConfigSheet(configSheet);
      prepareDirectory(activeSpreadsheet, configSheet);
      promptFillTemplate(ui);

    } else {
      ui.alert("Attention!", "This is not an empty spreadsheet.\nPlease create new spreadsheet.\n\nScript will abort.", ui.ButtonSet.OK);
    }
  }

}

/**
 * Check if still enough time to run another process. Appscript only allow 6 minutes of execution time.
 * @param {Date} initialTime Instance of `new Date()` from the initial execution time.
 * @param {Number} processDuration Estimated time (in seconds) for the process to complete
 */
function isEnoughTime(initialTime, processDuration) {
  let currentTime = new Date();
  let milisecondsDifference = currentTime.getTime() - initialTime.getTime();
  let secondsLeft = 360 - (milisecondsDifference / 1000);
  // Uncomment line below for more verbose output.
  // Logger.log('**-- ' + secondsLeft + ' seconds left --**');
  return secondsLeft > processDuration ? true : false;
}

/**
 * Script for onOpen() - generate sijil on selected row
 */
function actionGenerateGdoc() {
  let initialTime         = new Date();
  let confVar             = getConfigVariable();
  let generatedFolderGdoc = confVar.generatedFolderGdoc;
  let templateDoc         = confVar.templateDoc;
  let sheet               = confVar.dataSheet;
  let selectedRange       = sheet.getActiveRange();

  let forloopStart = selectedRange.getRowIndex();
  let forloopEnd   = forloopStart + selectedRange.getNumRows();
  for (let rowid = forloopStart; rowid < forloopEnd; rowid++) {
    if (isEnoughTime(initialTime, 25)) {
      let candidateData = new Array();
      let varTagList = sheet.getRange(2, 9, 1, sheet.getLastColumn()-8).getValues()[0];
      let varTagValueList = sheet.getRange(rowid, 9, 1, sheet.getLastColumn()-8).getValues()[0];
      varTagList.forEach((item, index) => { 
        candidateData.push({ varTag:item, varTagValue:varTagValueList[index] });
      });
      let candidateFile = generateGdocFile(generatedFolderGdoc, templateDoc, candidateData);
      sheet.getRange(rowid, 1).setValue(candidateFile.getUrl());
      sheet.getRange(rowid, 2).setValue(candidateFile.getId());
      Logger.log('Rowid: ' + rowid + ' - File created: ' + varTagValueList[0]);
    }
  }
}

/**
 * Script for onOpen() - convert sijil from Gdoc to PDF on selected row
 */
function actionGeneratePdf() {
  let initialTime   = new Date();
  let confVar       = getConfigVariable();
  let sheet         = confVar.dataSheet;
  let targetFolder  = confVar.generatedFolderPdf;
  let selectedRange = sheet.getActiveRange();

  let forloopStart = selectedRange.getRowIndex();
  let forloopEnd   = forloopStart + selectedRange.getNumRows();
  for (let rowid = forloopStart; rowid < forloopEnd; rowid++) {
    if (isEnoughTime(initialTime, 25)) {
      let gdocId       = sheet.getRange(rowid, 2).getValue();
      let gdocFile     = DriveApp.getFileById(gdocId);
      let convertedPdf = convertToPdf(gdocFile, targetFolder);
      sheet.getRange(rowid, 3).setValue(convertedPdf.getUrl());
      sheet.getRange(rowid, 4).setValue(convertedPdf.getId());
      Logger.log("File converted for rowid: " + rowid);
    }
  }
}

/**
 * Script for onOpen() - send email with PDF attachment on selected row
 */
function actionSendEmail() {
  let initialTime   = new Date();
  let confVar       = getConfigVariable();
  let sheet         = confVar.dataSheet;
  let selectedRange = sheet.getActiveRange();

  let forloopStart = selectedRange.getRowIndex();
  let forloopEnd = forloopStart + selectedRange.getNumRows();
  for (let rowid = forloopStart; rowid < forloopEnd; rowid++) {
    if (isEnoughTime(initialTime, 25)) {
      let recipientEmail = sheet.getRange(rowid, 5).getValue();
      let recipientAttachmentId = sheet.getRange(rowid, 4).getValue();
      sendEmailTo(confVar, recipientEmail, recipientAttachmentId);
      sheet.getRange(rowid, 6).setValue('DONE');
      Logger.log('Email sent to: ' + recipientEmail);
    }
  };
}

/**
 * Generate a single sijil as GoogleDoc object.
 * @param {Object} generatedFolderGdoc target folder for generated file
 * @param {Object} templateDoc GoogleDoc template file to copy from 
 * @param {Array} candidateData Array of candidateData objects
 */
function generateGdocFile(generatedFolderGdoc, templateDoc, candidateData) {
  let newDocObj = templateDoc.makeCopy(generatedFolderGdoc);
  let docApp = DocumentApp.openById(newDocObj.getId());
  let body = docApp.getBody();
  candidateData.forEach(item => {
    body.replaceText(item.varTag.toString(), item.varTagValue.toString());
  });
  newDocObj.setName(candidateData[0].varTagValue.toString());
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
 * Convert selection to UPPERCASE value.
 */
function toUpperCase() {
  let selectedRange = SpreadsheetApp.getActiveRange();
  let dataList = selectedRange.getValues();
  for (let i=0; i<dataList.length; i++) {
    for (let j=0; j<dataList[i].length; j++) {
      let value = dataList[i][j];
      if (!(value instanceof Date)) {
        value = value.toString().toUpperCase();
      };
    };
  };
  selectedRange.setValues(dataList);
}

/**
 * Prompt user about license.
 */
function aboutLicense() {
  let title = 'Open Source';
  let subtitle = `
    This app is open source and free to use under the terms of GNU General Public License v3.0.

    Copyright (C) 2025  akirasy
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
