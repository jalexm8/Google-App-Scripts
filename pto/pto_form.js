// Change these values.
const spreadsheetID = '';
const formID = '';
const calendarID = '';

/**
 * Transform the user responses to an array.
 *
 * @return {array} userResponses
 */
function formResponsesToArray() {
  const form = FormApp.getActiveForm();
  const formResponses = form.getResponses();
  const lastResponse =
    formResponses[formResponses.length - 1].getItemResponses();
  const userEmail =
    formResponses[formResponses.length - 1].getRespondentEmail();
  const startDate = lastResponse[0].getResponse();
  const endDate = lastResponse[1].getResponse();
  const formattedStartDate = new Date(startDate);
  const formattedEndDate = new Date(endDate);
  const userResponses = {
    'email': userEmail,
    'start date': formattedStartDate,
    'end date': formattedEndDate,
    'form start date': startDate,
    'form end date': endDate,
  };

  return userResponses;
}

/**
 * Perform some basic error checking.
 *
 * @param {array} userResponses
 * @param {date} dataSheet
 * @param {number} userRow
 *
 * @return {string} errorMsg
 */
function errorChecking(userResponses, dataSheet, userRow) {
  /* This function is where you would also add policy specific errors,
  *  such as maximum amount of PTO to take per request etc.
  */
  const PTOREMAININGCOLUMN = 5;

  let errorMsg = '';

  if ( userRow == null ) {
    errorMsg += 'ERROR: Email is not in system.';
    return errorMsg;
  }

  if ( userResponses['start date'].getTime() >
  userResponses['end date'].getTime() ) {
    errorMsg += 'ERROR: End date cannot be before start date.<br>';
  }

  const remainingPTO =
    dataSheet.getRange(userRow, PTOREMAININGCOLUMN).getValue();
  if ( remainingPTO <= 0 ) {
    errorMsg += 'ERROR: You do not have enough PTO remaining to request these' +
    ' dates.<br>';
  }

  const eventOnCal = checkCalendar(userResponses, dataSheet, userRow);
  if ( eventOnCal === true ) {
    errorMsg += 'ERROR: There is already a pending request during these ' +
    'dates.<br>';
  }

  return errorMsg;
}

/**
 * Get the row of the user.
 *
 * @param {date} dataSheet
 * @param {string} userEmail
 *
 * @return {number} i
 */
function getUserRow(dataSheet, userEmail) {
  const USEREMAILCOLUMN = 1; // this may not be right.

  const sheetValues = dataSheet.getDataRange().getValues();

  for ( let i = 0; i < sheetValues.length; i++) {
    if ( sheetValues[i][USEREMAILCOLUMN] == userEmail) {
      return i+1;
    }
  }
}

/**
 * Send the error email to user.
 *
 * @param {string} errorMsg
 * @param {string} userEmail
 */
function sendErrorEmail(errorMsg, userEmail) {
  MailApp.sendEmail({
    to: userEmail,
    subject: 'PTO request is not valid.',
    htmlBody: 'Please see below for possible error messages: <br><br>' +
      errorMsg,
  });
}

/**
 * Calculate how many days are between the start and end date.
 *
 * @param {date} startDate
 * @param {date} endDate
 *
 * @return {number} daysBetween
 */
function daysBetweenStartAndEndDate(startDate, endDate) {
  const timeBetween = endDate.getTime() - startDate.getTime();
  const daysBetween = timeBetween / (1000 * 3600 * 24);

  return daysBetween + 1;
}

/**
 * Add requested days to the spreadsheet.
 *
 * @param {string} dataSheet
 * @param {number} userRow
 * @param {number} ptoDaysRequested
 */
function addRequestedDaysToSpreadSheet(dataSheet, userRow, ptoDaysRequested) {
  const PTOREQUESTEDCOLUMN = 6;

  const prevRequestedTotal =
    dataSheet.getRange(userRow, PTOREQUESTEDCOLUMN).getValue();

  dataSheet.getRange(userRow, PTOREQUESTEDCOLUMN)
      .setValue(prevRequestedTotal + ptoDaysRequested);
}

/**
 * Send success email.
 *
 * @param {array} userResponses
 * @param {string} dataSheet
 * @param {number} userRow
 * @param {number} ptoDaysRequested
 */
function sendSuccessEmail(userResponses, dataSheet, userRow, ptoDaysRequested) {
  const PTOREMAININGCOLUMN = 5;

  const startDate = userResponses['start date'].toLocaleDateString();
  const endDate = userResponses['end date'].toLocaleDateString();
  const remainingPTO =
    dataSheet.getRange(userRow, PTOREMAININGCOLUMN).getValue();

  MailApp.sendEmail({
    to: userResponses['email'],
    subject: 'PTO request has been sent for approval.',
    htmlBody: 'Your request to take the below dates as PTO has been sent for ' +
      'approval: <br>' +
      startDate +
      ' - ' +
      endDate +
      '<br><br>' +
      'Total PTO requested: ' +
      ptoDaysRequested +
      '<br>' +
      'Remaining PTO: ' +
      remainingPTO,
  });
}

/**
 * Send the line manager an email.
 *
 * @param {array} userResponses
 * @param {string} dataSheet
 * @param {number} userRow
 * @param {number} ptoDaysRequested
 */
function sendLineManagerEmail(userResponses,
    dataSheet, userRow, ptoDaysRequested) {
  const LINEMANAGEREMAIL = 3;
  const startDate = userResponses['start date'].toLocaleDateString();
  const endDate = userResponses['end date'].toLocaleDateString();
  const userLineManagerEmail =
    dataSheet.getRange(userRow, LINEMANAGEREMAIL).getValue();

  MailApp.sendEmail({
    to: userLineManagerEmail,
    subject: 'PTO request for ' + userResponses['email'],
    htmlBody: userResponses['email'] +
      ' has requested your approval for the below dates: <br>' +
      startDate +
      ' - ' +
      endDate +
      '<br><br>' +
      'Total PTO: ' +
      ptoDaysRequested +
      '<br><br>' +
      'Follow the below link to approve:<br>' +
      'https://docs.google.com/forms/d/e/' +
      formID +
      '/viewform?usp=pp_url&entry.1215773484=' +
      userResponses['email'] +
      '&entry.2132041378=' +
      userResponses['form start date'] +
      '&entry.1828168590=' +
      userResponses['form end date'] +
      '&entry.1716043255=Approved',
  });
}

/**
 * Add pending PTO to calendar.
 *
 * @param {array} userResponses
 * @param {string} dataSheet
 * @param {number} userRow
 */
function addPendingPtoToCal(userResponses, dataSheet, userRow) {
  const USERNAMECOLUMN = 1;
  const LINEMANAGEREMAIL = 3;
  const userName = dataSheet.getRange(userRow, USERNAMECOLUMN).getValue();
  const googleCal = CalendarApp.getCalendarById(calendarID);
  const day = 60 * 60 * 24 * 1000;
  const endDatePlus1 = new Date(userResponses['end date'].getTime() + day);
  const userLineManagerEmail =
    dataSheet.getRange(userRow, LINEMANAGEREMAIL).getValue();
  const invitees = userResponses['email'] + ',' + userLineManagerEmail;

  googleCal.createAllDayEvent(userName + ' - [PENDING]',
      userResponses['start date'],
      endDatePlus1,
      {guests: invitees, sendInvites: true});
}

/**
 * Check calendar for event.
 *
 * @param {array} userResponses
 * @param {string} dataSheet
 * @param {number} userRow
 *
 * @return {boolean} eventOnCal
 */
function checkCalendar(userResponses, dataSheet, userRow) {
  const USERNAMECOLUMN = 1;
  const userName = dataSheet.getRange(userRow, USERNAMECOLUMN).getValue();
  const googleCal = CalendarApp.getCalendarById(calendarID);
  const day = 60 * 60 * 24 * 1000;
  const endDatePlus1 = new Date(userResponses['end date'].getTime() + day);
  const pendingCalEvents =
    googleCal.getEvents(userResponses['start date'],
        endDatePlus1, {search: userName + '[PENDING]'});
  const approvedCalEvents =
    googleCal.getEvents(userResponses['start date'],
        endDatePlus1, {search: userName + '[APPROVED]'});
  let eventOnCal = false;

  if ( pendingCalEvents.length > 0 || approvedCalEvents.length > 0) {
    eventOnCal = true;
  }

  return eventOnCal;
}

/**
 * Add form to the spreadsheet.
 *
 * @param {array} userResponses
 * @param {string} spreadSheet
 * @param {string} dataSheet
 * @param {number} userRow
 */
function addFormToSheet(userResponses, spreadSheet, dataSheet, userRow) {
  const USERNAMECOLUMN = 1;
  const USEREMAILCOLUMN = 2;
  const LINEMANAGEREMAIL = 3;
  const STARTDATECOLUMN = 4;
  const ENDDATECOLUMN = 5;
  const APPROVALFORMCOLUMN = 6;
  const DAYSNOTAPPROVEDCOLUMN = 7;
  const pendingFormSheet =
    spreadSheet.getSheetByName('pendingPtoApprovalForms');
  const lastRow = pendingFormSheet.getLastRow() + 1;
  const approvalForm = 'https://docs.google.com/forms/d/e/' +
    formID +
    '/viewform?usp=pp_url&entry.1215773484=' +
    userResponses['email'] +
    '&entry.2132041378=' +
    userResponses['form start date'] +
    '&entry.1828168590=' +
    userResponses['form end date'] +
    '&entry.1716043255=Approved';
  const userLineManagerEmail =
    dataSheet.getRange(userRow, LINEMANAGEREMAIL).getValue();
  const userName = dataSheet.getRange(userRow, USERNAMECOLUMN).getValue();

  pendingFormSheet.getRange(lastRow, USERNAMECOLUMN).setValue(userName);
  pendingFormSheet.getRange(lastRow, USEREMAILCOLUMN)
      .setValue(userResponses['email']);
  pendingFormSheet.getRange(lastRow, LINEMANAGEREMAIL)
      .setValue(userLineManagerEmail);
  pendingFormSheet.getRange(lastRow, STARTDATECOLUMN)
      .setValue(String(userResponses['form start date']));
  pendingFormSheet.getRange(lastRow, ENDDATECOLUMN)
      .setValue(String(userResponses['form end date']));
  pendingFormSheet.getRange(lastRow, APPROVALFORMCOLUMN).setValue(approvalForm);
  pendingFormSheet.getRange(lastRow, DAYSNOTAPPROVEDCOLUMN).setValue(0);
}

/**
 * Get's data from the form submit.
 *
 * @param {array} e
 */
// eslint-disable-next-line no-unused-vars, require-jsdoc
function onFormSubmit(e) {
  const userResponses = formResponsesToArray();
  const spreadSheet = SpreadsheetApp.openById(spreadsheetID);
  const dataSheet = spreadSheet.getSheetByName('data');
  const userRow = getUserRow(dataSheet, userResponses['email']);
  const errorMsg = errorChecking(userResponses, dataSheet, userRow);

  if ( errorMsg.length > 0 ) {
    sendErrorEmail(errorMsg, userResponses['email']);
  } else {
    const ptoDaysRequested =
      daysBetweenStartAndEndDate(userResponses['start date'],
          userResponses['end date']);
    addRequestedDaysToSpreadSheet(dataSheet, userRow, ptoDaysRequested);
    sendLineManagerEmail(userResponses, dataSheet, userRow, ptoDaysRequested);
    sendSuccessEmail(userResponses, dataSheet, userRow, ptoDaysRequested);
    addPendingPtoToCal(userResponses, dataSheet, userRow);
    addFormToSheet(userResponses, spreadSheet, dataSheet, userRow);
  }
}
