function formResponsesToArray() {
  var form = FormApp.getActiveForm();
  var formResponses = form.getResponses();
  var lastResponse = formResponses[formResponses.length - 1].getItemResponses();
  var userEmail = formResponses[formResponses.length - 1].getRespondentEmail();
  var startDate = lastResponse[0].getResponse();
  var endDate = lastResponse[1].getResponse();
  var formattedStartDate = new Date(startDate);
  var formattedEndDate = new Date(endDate);

  var userResponses = {
    'email': userEmail,
    'start date': formattedStartDate,
    'end date': formattedEndDate,
    'form start date': startDate,
    'form end date': endDate,
  };

  return userResponses;
}

function errorChecking(startDate, endDate, dataSheet, userRow) {
  /* This function is where you would also add policy specific errors,
  *  such as maximum amount of PTO to take per request etc.
  */
  const PTOREMAININGCOLUMN = 5;

  var errorMsg = "";

  if ( userRow == null ) {
    errorMsg += "ERROR: Email is not in system."
    return errorMsg;
  }

  if ( startDate.getTime() > endDate.getTime() ) {
    errorMsg += "ERROR: End date cannot be before start date.<br>";
  }

  var remainingPTO = dataSheet.getRange(userRow, PTOREMAININGCOLUMN).getValue();
  if ( remainingPTO <= 0 ) {
    errorMsg += "ERROR: You do not have enough PTO remaining to request these dates.<br>";
  }

  return errorMsg;
}

function getUserRow(dataSheet, userEmail) {
  const USEREMAILCOLUMN = 1; //this may not be right.

  var sheetValues = dataSheet.getDataRange().getValues();

  for ( var i = 0; i < sheetValues.length ; i++) {
    if ( sheetValues[i][USEREMAILCOLUMN] == userEmail) {
      return i+1;
    }
  }
}

function sendErrorEmail(errorMsg, userEmail) {
  MailApp.sendEmail({
    to: userEmail,
    subject: "PTO request is not valid.",
    htmlBody: "Please see below for possible error messages: <br><br>" + errorMsg
  });
}

function daysBetweenStartAndEndDate(startDate, endDate) {
  var timeBetween = endDate.getTime() - startDate.getTime();
  var daysBetween = timeBetween / (1000 * 3600 * 24);

  return daysBetween + 1;
}

function addRequestedDaysToSpreadSheet(dataSheet, userRow, ptoDaysRequested) {
  const PTOREQUESTEDCOLUMN = 6;

  var prevRequestedTotal = dataSheet.getRange(userRow, PTOREQUESTEDCOLUMN).getValue();
  dataSheet.getRange(userRow, PTOREQUESTEDCOLUMN).setValue(prevRequestedTotal + ptoDaysRequested);
}

function sendSuccessEmail(userResponses, dataSheet, userRow, ptoDaysRequested) {
  const PTOREMAININGCOLUMN = 5;

  var startDate = userResponses['start date'].toLocaleDateString();
  var endDate = userResponses['end date'].toLocaleDateString();
  var remainingPTO = dataSheet.getRange(userRow, PTOREMAININGCOLUMN).getValue();

  MailApp.sendEmail({
    to: userResponses['email'],
    subject: "PTO request has been sent for approval.",
    htmlBody: "Your request to take the below dates as PTO has been sent for approval: <br>" +
              startDate + " - " + endDate + "<br><br>" +
              "Total PTO requested: " + ptoDaysRequested + "<br>" +
              "Remaining PTO: " + remainingPTO
  });
}

function sendLineManagerEmail(userResponses, dataSheet, userRow, ptoDaysRequested) {
  const LINEMANAGEREMAIL = 3;

  var startDate = userResponses['start date'].toLocaleDateString();
  var endDate = userResponses['end date'].toLocaleDateString();
  var userLineManagerEmail = dataSheet.getRange(userRow, LINEMANAGEREMAIL).getValue();

  MailApp.sendEmail({
    to: userLineManagerEmail,
    subject: "PTO request for " + userResponses['email'],
    htmlBody: userResponses['email'] + " has requested your approval for the below dates: <br>" +
              startDate + " - " + endDate + "<br><br>" + 
              "Total PTO: " + ptoDaysRequested + "<br>" +
              "<br>" +
              "Follow the below link to approve:<br>" +
              "https://docs.google.com/forms/d/e/<FORM_ID>/viewform?usp=pp_url&entry.1215773484=" + userResponses['email'] + "&entry.2132041378=" + userResponses['form start date'] + "&entry.1828168590=" + userResponses['form end date'] + "&entry.1716043255=Approved"
  });
}

function onFormSubmit(e) {
  var userResponses = formResponsesToArray();
  var spreadSheet = SpreadsheetApp.openById("<SPREADSHEET_ID>");
  var dataSheet = spreadSheet.getSheetByName('data');
  var userRow = getUserRow(dataSheet, userResponses['email']);
  var errorMsg = errorChecking(userResponses['start date'], userResponses['end date'], dataSheet, userRow);

  if ( errorMsg.length > 0 ) {
    sendErrorEmail(errorMsg, userResponses['email']);
  } else {
    var ptoDaysRequested = daysBetweenStartAndEndDate(userResponses['start date'], userResponses['end date']);
    addRequestedDaysToSpreadSheet(dataSheet, userRow, ptoDaysRequested);
    sendLineManagerEmail(userResponses, dataSheet, userRow, ptoDaysRequested);
    sendSuccessEmail(userResponses, dataSheet, userRow, ptoDaysRequested);
  }
}
