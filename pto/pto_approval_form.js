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
  const lineManagerEmail =
    formResponses[formResponses.length - 1].getRespondentEmail();
  const userEmail = lastResponse[0].getResponse();
  const startDate = lastResponse[1].getResponse();
  const endDate = lastResponse[2].getResponse();
  const requestResponse = lastResponse[3].getResponse();
  const declineReason = lastResponse[4].getResponse();
  const formattedStartDate = new Date(startDate);
  const formattedEndDate = new Date(endDate);
  const userResponses = {
    'line manager email': lineManagerEmail,
    'user email': userEmail,
    'start date': formattedStartDate,
    'end date': formattedEndDate,
    'request response': requestResponse,
    'decline reason': declineReason,
    'form start date': startDate,
    'form end date': endDate,
  };

  return userResponses;
}

/**
 * Perform some basic error checking.
 *
 * @param {date} startDate
 * @param {date} endDate
 * @param {date} dataSheet
 * @param {number} userRow
 *
 * @return {string} errorMsg
 */
function errorChecking(startDate, endDate, dataSheet, userRow) {
  /* This function is where you would also add policy specific errors,
  *  such as maximum amount of PTO to take per request etc.
  */
  const PTOREMAININGCOLUMN = 5;
  let errorMsg = '';

  if ( userRow == null ) {
    errorMsg += 'ERROR: User email is not in system.';
    return errorMsg;
  }

  if ( startDate.getTime() > endDate.getTime() ) {
    errorMsg += 'ERROR: End date cannot be before start date.<br>';
  }

  const remainingPTO =
    dataSheet.getRange(userRow, PTOREMAININGCOLUMN).getValue();
  if ( remainingPTO <= 0 ) {
    errorMsg += 'ERROR: The user does not have enough PTO remaining to ' +
    'request these dates.<br>';
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
 * Send the line manager an error email.
 *
 * @param {string} errorMsg
 * @param {string} lineManagerEmail
 */
function sendErrorEmail(errorMsg, lineManagerEmail) {
  MailApp.sendEmail({
    to: lineManagerEmail,
    subject: 'PTO approval is not valid.',
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
 * Remove requested amount of days from the user's amount.
 *
 * @param {string} dataSheet
 * @param {number} userRow
 * @param {number} ptoDaysRequested
 */
function removeRequestedDaysFromSpreadsheet(dataSheet,
    userRow, ptoDaysRequested) {
  const PTOREQUESTEDCOLUMN = 6;
  const prevRequestedTotal =
    dataSheet.getRange(userRow, PTOREQUESTEDCOLUMN).getValue();

  dataSheet.getRange(userRow, PTOREQUESTEDCOLUMN)
      .setValue(prevRequestedTotal - ptoDaysRequested);
}

/**
 * Send the user an email on if leave was approved/denied.
 *
 * @param {array} userResponses
 * @param {string} dataSheet
 * @param {number} userRow
 * @param {number} ptoDaysRequested
 */
function sendUserEmail(userResponses, dataSheet, userRow, ptoDaysRequested) {
  const PTOREMAININGCOLUMN = 5;
  const startDate = userResponses['start date'].toLocaleDateString();
  const endDate = userResponses['end date'].toLocaleDateString();
  const remainingPTO =
    dataSheet.getRange(userRow, PTOREMAININGCOLUMN).getValue();
  const declineReason =
    userResponses['decline reason'] || 'No reason provided.';

  if ( userResponses['request response'] == 'Declined') {
    MailApp.sendEmail({
      to: userResponses['user email'],
      subject: 'Your PTO request has been declined.',
      htmlBody: 'Your request to take the below dates as PTO ' +
        'has been declined: <br>' +
        startDate +
        ' - ' +
        endDate +
        '<br><br>' +
        'Total PTO requested: ' +
        ptoDaysRequested +
        '<br>' +
        'Decline reason: ' +
        declineReason +
        '<br><br>' +
        'Remaining PTO: ' +
        remainingPTO,
    });
  } else {
    MailApp.sendEmail({
      to: userResponses['user email'],
      subject: 'PTO request has been approved.',
      htmlBody: 'Your request to take the below dates as PTO ' +
        'has been approved: <br>' +
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
}

/**
 * Send the user's line manager an email on if leave was approved/denied.
 *
 * @param {array} userResponses
 * @param {number} ptoDaysRequested
 */
function sendLineManagerEmail(userResponses, ptoDaysRequested) {
  const startDate = userResponses['start date'].toLocaleDateString();
  const endDate = userResponses['end date'].toLocaleDateString();
  const declineReason =
    userResponses['decline reason'] || 'No reason provided.';

  if ( userResponses['request response'] == 'Declined') {
    MailApp.sendEmail({
      to: userResponses['line manager email'],
      subject: userResponses['user email'] +
        '\'s PTO request has been declined.',
      htmlBody: 'You have declined ' +
        userResponses['user email'] +
        '\'s request to take the below dates as PTO: <br>' +
        startDate +
        ' - ' +
        endDate +
        '<br><br>' +
        'Total PTO requested: ' +
        ptoDaysRequested +
        '<br>' +
        'Decline reason: ' +
        declineReason,
    });
  } else {
    MailApp.sendEmail({
      to: userResponses['line manager email'],
      subject: userResponses['user email'] +
        '\'s PTO request has been approved.',
      htmlBody: 'You have approved ' +
        userResponses['user email'] +
        '\'s request to take the below dates as PTO: <br>' +
        startDate +
        ' - ' +
        endDate +
        '<br><br>' +
        'Total PTO requested: ' +
        ptoDaysRequested,
    });
  }
}

/**
 * Add the requested days to the total.
 *
 * @param {string} dataSheet
 * @param {number} userRow
 * @param {number} ptoDaysRequested
 */
function addApprovedDaysToSpreadSheet(dataSheet, userRow, ptoDaysRequested) {
  const PTOREQUESTEDCOLUMN = 7;
  const prevRequestedTotal =
    dataSheet.getRange(userRow, PTOREQUESTEDCOLUMN).getValue();

  dataSheet.getRange(userRow, PTOREQUESTEDCOLUMN)
      .setValue(prevRequestedTotal + ptoDaysRequested);
}

/**
 * Update the calendar with an event for the PTO.
 *
 * @param {array} userResponses
 * @param {string} dataSheet
 * @param {number} userRow
 */
function updateCalender(userResponses, dataSheet, userRow) {
  const USERNAMECOLUMN = 1;
  const userName = dataSheet.getRange(userRow, USERNAMECOLUMN).getValue();
  const ptoCalendar = CalendarApp.getCalendarById(calendarID);
  const day = 60 * 60 * 24 * 1000;
  const endDatePlus1 = new Date(userResponses['end date'].getTime() + day);
  const prevCalEvents = ptoCalendar.getEvents(userResponses['start date'],
      endDatePlus1, {search: userName + '[PENDING]'});

  if ( userResponses['request response'] == 'Approved') {
    for ( let i = 0; i < prevCalEvents.length; i++) {
      prevCalEvents[i].setTitle(userName + ' - [APPROVED]');
    }
  } else {
    for ( let i = 0; i < prevCalEvents.length; i++) {
      prevCalEvents[i].deleteEvent();
    }
  }
}

/**
 * Delete request row from pending form sheet.
 *
 * @param {array} userResponses
 * @param {string} spreadSheet
 */
function deleteRequestRow(userResponses, spreadSheet) {
  const pendingFormSheet =
    spreadSheet.getSheetByName('pendingPtoApprovalForms');
  const requestRow = getRequestRow(userResponses, pendingFormSheet);

  pendingFormSheet.deleteRow(requestRow);
}

/**
 * Get the request row
 *
 * @param {array} userResponses
 * @param {string} pendingFormSheet
 *
 * @return {number} i
 */
function getRequestRow(userResponses, pendingFormSheet) {
  const APPROVALEMAILCOLUMN = 5;

  const sheetValues = pendingFormSheet.getDataRange().getValues();
  const approvalForm = 'https://docs.google.com/forms/d/e/' +
    formID +
    '/viewform?usp=pp_url&entry.1215773484=' +
    userResponses['user email'] +
    '&entry.2132041378=' +
    userResponses['form start date'] +
    '&entry.1828168590=' +
    userResponses['form end date'] +
    '&entry.1716043255=Approved';

  for ( let i = 0; i < sheetValues.length; i++) {
    if ( sheetValues[i][APPROVALEMAILCOLUMN] == approvalForm) {
      return i+1;
    }
  }
}

/**
 * Get data once the form has been submitted.
 *
 * @param {array} e
 */
// eslint-disable-next-line no-unused-vars, require-jsdoc
function onFormSubmit(e) {
  const userResponses = formResponsesToArray();
  const spreadSheet = SpreadsheetApp.openById(spreadsheetID);
  const dataSheet = spreadSheet.getSheetByName('data');
  const userRow = getUserRow(dataSheet, userResponses['user email']);
  const errorMsg = errorChecking(userResponses['start date'],
      userResponses['end date'], dataSheet, userRow);

  if ( errorMsg.length > 0 ) {
    sendErrorEmail(errorMsg, userResponses['line manager email']);
  } else {
    const ptoDaysRequested =
      daysBetweenStartAndEndDate(userResponses['start date'],
          userResponses['end date']);

    removeRequestedDaysFromSpreadsheet(dataSheet, userRow, ptoDaysRequested);
    if ( userResponses['request response'] == 'Approved') {
      addApprovedDaysToSpreadSheet(dataSheet, userRow, ptoDaysRequested);
    }

    sendLineManagerEmail(userResponses, ptoDaysRequested);
    sendUserEmail(userResponses, dataSheet, userRow, ptoDaysRequested);
    updateCalender(userResponses, dataSheet, userRow);
    deleteRequestRow(userResponses, spreadSheet);
  }
}
