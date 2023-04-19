function onNewRow(e) {
    // Read data from the new row in the Google Sheet
    const data = readDataFromSheet(e);
  
    // Generate the name of the PDF to be created
    const pdfName = generatePdfName(data);
  
    // Calculate the end date and format all dates as dd-mm-yyyy
    const formattedDates = formatDates(data);
  
    // Choose the correct template based on the true/false statements
    const templateId = chooseTemplate(data);
  
    // Populate the template with the data
    const populatedSlideId = populateTemplate(templateId, data, formattedDates);
  
    // Export the populated Google Slides file as a PDF
    const pdfFile = exportSlideAsPdf(populatedSlideId, pdfName);
  
    // Generate the email subject and body (HTML)
    const emailSubject = generateEmailSubject(data);
    const emailBodyHtml = generateEmailBodyHtml(data);
  
    // Send the email with the PDF attached using Gmail
    sendEmailWithPdf(data.email, emailSubject, emailBodyHtml, pdfFile);
}

function readDataFromSheet(e) {
    const sheet = e.source.getSheetByName('YourSheetName');
    const row = e.range.getRow();
    const data = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // Convert the row data into an object for easier manipulation
    const rowData = {
      name: data[0],
      email: data[1],
      // Add other relevant fields from your Google Sheet
    };
    
    return rowData;
}
  

function generatePdfName(data) {
    // Replace the logic below with your desired PDF naming convention
    return data.name + "_Contract_" + new Date().toISOString().substring(0, 10);
}  

function formatDates(data) {
    const startDate = new Date(data.startDate);
    const endDate = new Date(startDate);
    endDate.setMonth(endDate.getMonth() + data.contractMonths);
    endDate.setDate(endDate.getDate() - 1);
  
    return {
      startDate: formatDate(startDate),
      endDate: formatDate(endDate),
    };
}
  
function formatDate(date) {
    const day = date.getDate();
    const month = date.getMonth() + 1;
    const year = date.getFullYear();
    return `${("0" + day).slice(-2)}-${("0" + month).slice(-2)}-${year}`;
}  

function chooseTemplate(data) {
  // Replace the logic below with your specific true/false conditions
  let templateId;
  if (data.condition1) {
    templateId = 'templateId1';
  } else if (data.condition2) {
    templateId = 'templateId2';
  } else {
    templateId = 'templateId3';
  }
  return templateId;
}

function populateTemplate(templateId, data, formattedDates) {
  const slide = SlidesApp.openById(templateId);
  const fileId = DriveApp.getFileById(templateId).makeCopy().getId();
  const newSlide = SlidesApp.openById(fileId);

  // Replace placeholders with actual data
  newSlide.replaceAllText('{{name}}', data.name);
  newSlide.replaceAllText('{{startDate}}', formattedDates.startDate);
  newSlide.replaceAllText('{{endDate}}', formattedDates.endDate);
  // Replace other placeholders as needed
  
  // Save and return the new slide ID
  newSlide.saveAndClose();
  return fileId;
}


function exportSlideAsPdf(populatedSlideId, pdfName) {
  const url = `https://docs.google.com/presentation/d/${populatedSlideId}/export/pdf`;
  const options = {
    method: 'get',
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
  };
  const response = UrlFetchApp.fetch(url, options);
  return DriveApp.createFile(response.getBlob().setName(pdfName));
}


function generateEmailSubject(data) {
  // Replace with your desired email subject logic
  return `Your Contract for ${data.name}`;
}

function generateEmailBodyHtml(data) {
  // ...
}

function sendEmailWithPdf(email, emailSubject, emailBodyHtml, pdfFile) {
  // ...
}

// Set up the trigger to run the script when a new row is added
function createTrigger() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet();
    ScriptApp.newTrigger('onNewRow')
      .forSpreadsheet(sheet)
      .onEdit()
      .create();
}