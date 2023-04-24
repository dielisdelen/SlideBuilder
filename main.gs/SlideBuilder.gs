const templateMap = {
  'TTTgeen': '1T1aW3Cfq25c6yNlkrlfezCmM06Yk_ve_puJ6v2gY7X4',
  'TTTduurzaam': '1ehkFXwqi6yrQ4bnxDc_3u_a5zJ9w7C-j0ncQ5OsT6cY',
  'TTTnormaal': '1yo4ChlHRCA3bRkQzbPT7LLR9BLgrfGu3iY3p0jnqnnE',
  'TTFgeen': '1zvz4nIV-oJDpETKX8EZxQd9DgNiD_dw4rGsOeI0aF4A',
  'TTFduurzaam': '11-QODS910m2vDf98rlV4u28ciFRNpgTzb2JXk2dT_zM',
  'TTFnormaal': '1bfOlE1wKsXxWQ3VIKP_l96ANTULHI3LEfc7jWFJIGqA',
  'TFTgeen': '1aLJtNXcYl0jKqcGqptwqhNow1GIB00Wk2Y8P4L71jno',
  'TFTduurzaam': '1A_82FCrFGtpxc3TxlZ-530wJ6FiXP97CqD_4v111VC4',
  'TFTnormaal': '1knVeAvlFLsgZuWGR8q0N5yTe-ioNF2kgJ8hFgzR9PEM',
  'FTTgeen': '1wm2_MlGinkqTpx-XQNR0YUX6Y7v8LraXtEZ4hHtS-pg',
  'FTTduurzaam': '1xnBMjWdvIygQTDc1rqDaF6UKv5rlY8BnIPzbe1wt6uE',
  'FTTnormaal': '1MyND4kO4YhWIgB3-5vd0hkgHF8gBJttMh9pm-5hxOYM',
  'FFTgeen': '15UuJcXG6shsUXOQx_IyHhQmTyZzcZmF0KIGMR803JjU',
  'FFTduurzaam': '1OjrUXmffnF5MOo1c9W-Slyzj2CQd_Sxsq4tIL1EGAiw',
  'FFTnormaal': '1P6VfaDB7bEm36nA60peT5f46E8RLAobsDJAKRBdzGRw',
  'FFFgeen': '1t4G2bOFG0FOjBnTyXMneDl1eb83JGT7XBM_ZG4S2D2Y',
  'FFFduurzaam': '1dp97RchECbVeBMRgo0ZFSCRweD_KzRE9zLFuK2dsSL8',
  'FFFnormaal': '1j3u8JtfxMV2q6ShgilVoUdjTNkjUIKNwUkQXdY-LpnE',
  'FTFgeen': '1yJopAGuZfnUjkk8eePe6gyEE_lX5IZaDabOFLy0bvJw',
  'FTFduurzaam': '12cBgS3mW2b5HF90ufSzSVRxtgvMFQTNstMcVBEVWC6A',
  'FTFnormaal': '/1JfK-L5KIxKuQKXIignU9hoHgUREZSm-U2I634W6qyn8',
  'TFFgeen': '1oq0kUu1UX136JtafnRMrQV9Kt36Vcj-H9SkINg90FSU',
  'TFFduurzaam': '1NmCvh5VysXBgZZQw9hAi3b8pOiLeEfzOejPxQwc29is',
  'TFFnormaal': '1Y6Iq8GzB8AIPNPVHG9NYocrFmXyETnEjzX6_Dr-iH3g'
};

function onNewRow() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('groene overeenkomst wizard');
  const lastRow = sheet.getLastRow();
  const data = sheet.getRange(lastRow, 1, 1, sheet.getLastColumn()).getValues()[0];

  // Convert the row data into an object for easier manipulation
  const rowData = {
    nameUser: data[6],
    emailUser: data[7],
    organisationUser: data[8],
    organisationAddress: data[9],
    employeeName: data[10],
    employeeAddress: data[11],
    employeeFunction: data[12],
    employeeCompensation: data[13],
    startContract: data[14],
    termContract: data[15],
    monthsContract: data[16],
    trialContract: data[17],
    hoursContract: data[18],
    holidayMoneyContract: data[19],
    holidayDaysContract: data[20],
    bonusContract: data[21],
    lunchContract: data[22],
    caoContract: data[23],
    caoNameContract: data[24],
    pensionContract: data[25],
    pensionNameContract: data[26],
  };

  // Log the rowData for testing purposes
  console.log(JSON.stringify(rowData));

  // Test generatePdfName function
  const pdfName = generatePdfName(rowData);
  console.log('Generated PDF Name: ', pdfName);

  // Test calculateEndDate function
  const endDate = calculateEndDate(rowData);
  console.log('Calculated End Date: ', endDate);

  // Test formatStartDate
  const newStartDate = formatStartDate(rowData.startContract);
  console.log('Format of the new startDate: ', newStartDate);

  // Test getTemplateFileId function
  const pensionContractShortName = getPensionContractShortName(rowData.pensionContract);
  const identifier =
    (rowData.bonusContract ? 'T' : 'F') +
    (rowData.lunchContract ? 'T' : 'F') +
    (rowData.caoContract ? 'T' : 'F') +
    pensionContractShortName;
  console.log('Generated Identifier: ', identifier);
  console.log('Row Data:', JSON.stringify(rowData));

  const templateFileId = getTemplateFileId(
    rowData.bonusContract,
    rowData.lunchContract,
    rowData.caoContract,
    pensionContractShortName
  );
  console.log('Template File ID: ', templateFileId);

  // Fill the template with data.
  const newPresentationId = fillTemplateWithDataV2(templateFileId, newStartDate, endDate, rowData);

  // Proceed with the next steps: create a PDF and send an email.
  const emailSubject = pdfName;
  const nameUser = rowData.nameUser;
  const emailBody = `<p>Beste ${nameUser},</p><p>In de bijlage vind je de groene arbeidsovereenkomst van ${rowData.employeeName}. Je kunt deze digitaal ondertekenen, bijvoorbeeld met <a href='https://signrequest.com/#/'>Signrequest</a>.</p><p>Daarna moeten de afspraken in de overeenkomst ook geïmplementeerd worden. Op onze website vind je daarvoor een <a href='http://groeneovereenkomst.nl/stappenplan'>stappenplan</a>.</p><p>Groeten,<br>Savannah Koomen</p>`;
  const recipientEmail = rowData.emailUser;
  sendEmailWithPdf(newPresentationId, emailSubject, emailBody, recipientEmail);

  // Post-process (delete row and slides file)
  const rowIndex = 2;
  postProcess(sheet, rowIndex, newPresentationId);

}

function createTrigger() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  ScriptApp.newTrigger('onNewRow')
    .forSpreadsheet(sheet)
    .onEdit()
    .create();
}

function generatePdfName(data) {
  return "Groene overeenkomst van " + data.employeeName;
}

function calculateEndDate(data) {
  if (data.termContract === 'bepaalde tijd') {
    const startDate = new Date(data.startContract);
    startDate.setMonth(startDate.getMonth() + data.monthsContract);
    startDate.setDate(startDate.getDate() - 1);

    const endDate = startDate.toLocaleDateString('nl-NL', {
      day: '2-digit',
      month: '2-digit',
      year: 'numeric',
    });

    return endDate;
  } else if (data.termContract === 'onbepaalde tijd') {
    return 'onbepaalde tijd';
  }

  return null;
}

function formatStartDate(dateString) {
  const startDate = new Date(dateString);
  const formattedStartDate = startDate.toLocaleDateString('nl-NL', {
    day: '2-digit',
    month: '2-digit',
    year: 'numeric',
  });

  return formattedStartDate;
}

function getTemplateFileId(bonusContract, lunchContract, caoContract, pensionContract) {
  const identifier =
    (bonusContract ? 'T' : 'F') +
    (lunchContract ? 'T' : 'F') +
    (caoContract ? 'T' : 'F') +
    pensionContract;

  return templateMap[identifier];
}


function getPensionContractShortName(pensionContract) {
  switch (pensionContract) {
    case 'de werkgever is aangesloten bij een duurzaam pensioenfonds.':
      return 'duurzaam';
    case 'de werkgever is aangesloten bij een \'normaal\' pensioenfonds':
      return 'normaal';
    case 'geen pensioen':
      return 'geen';
    default:
      console.error('Unknown pensionContract value:', pensionContract);
      return '';
  }
}

function fillTemplateWithDataV2(templateFileId, newStartDate, endDate, rowData) {
  const copyTitle = "Groene overeenkomst van " + rowData.employeeName;
  let copyFile = {
    title: copyTitle,
    parents: [{ id: "root" }],
  };
  copyFile = DriveApp.getFileById(templateFileId).makeCopy(copyTitle);
  const presentationCopyId = copyFile.getId();

  const requests = [
    { key: '{{werkgever}}', value: rowData.organisationUser },
    { key: '{{Adres werkgever}}', value: rowData.organisationAddress },
    { key: '{{Naam werknemer}}', value: rowData.employeeName },
    { key: '{{Adres werknemer}}', value: rowData.employeeAddress },
    { key: '{{Rol}}', value: rowData.employeeFunction },
    { key: '{{Salaris}}', value: String(rowData.employeeCompensation) },
    { key: '{{Start datum}}', value: newStartDate },
    { key: '{{Eind datum contract}}', value: endDate },
    { key: '{{Proeftijd}}', value: rowData.trialContract },
    { key: '{{Aantal uur}}', value: String(rowData.hoursContract) },
    { key: '{{Vakantiegeld}}', value: String(rowData.holidayMoneyContract) },
    { key: '{{Vakantiedagen}}', value: String(rowData.holidayDaysContract) },
    { key: '{{Pensioenfonds}}', value: rowData.pensionNameContract },
    { key: '{{Naam Werknemer}}', value: rowData.employeeName },
    { key: '{{CAO}}', value: rowData.caoNameContract },
  ].map(replacement => ({
    replaceAllText: {
      containsText: {
        text: replacement.key,
        matchCase: true,
      },
      replaceText: replacement.value,
    },
  }));

  const result = Slides.Presentations.batchUpdate(
    {
      requests: requests,
    },
    presentationCopyId
  );

  let numReplacements = 0;
  result.replies.forEach(function (reply) {
    numReplacements += reply.replaceAllText.occurrencesChanged;
  });

  console.log("Created presentation with ID: %s", presentationCopyId);
  console.log("Replaced %s text instances", numReplacements);

  return presentationCopyId;
}

function sendEmailWithPdf(newPresentationId, emailSubject, emailBody, recipientEmail) {
  // Export the presentation as a PDF
  const url = `https://docs.google.com/presentation/d/${newPresentationId}/export/pdf`;
  const options = {
    method: 'get',
    headers: {
      Authorization: 'Bearer ' + ScriptApp.getOAuthToken()
    }
  };
  const response = UrlFetchApp.fetch(url, options);
  const pdfBlob = response.getBlob().setName(emailSubject + '.pdf');

  // Send the email with the PDF attachment
  GmailApp.sendEmail(
    recipientEmail,
    emailSubject,
    emailBody,
    {
      attachments: [pdfBlob],
      htmlBody: emailBody
    }
  );
}

function deleteRow(sheet, rowIndex) {
  sheet.deleteRow(rowIndex);
}

function deleteSlidesFile(fileId) {
  DriveApp.getFileById(fileId).setTrashed(true);
}

function postProcess(sheet, rowIndex, fileId) {
  // Delete the Google Sheets row
  deleteRow(sheet, rowIndex);

  // Delete the Google Slides file
  deleteSlidesFile(fileId);
}
