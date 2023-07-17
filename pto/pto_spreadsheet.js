// Change this value
const spreadSheetID = '';

const PTOFORMS_USERNAME = 1;
const PTOFORMS_LINEMANAGERCOLUMN = 3;
const PTOFORMS_APPROVALFORMCOLUMN = 6;
const PTOFORMS_DAYSNOTAPPROVEDCOLUMN = 7;
const ptoFromSheetName = 'pendingPtoApprovalForms';
const spreadSheet = SpreadsheetApp.openById(spreadSheetID);
const ptoFormsSheet = spreadSheet.getSheetByName(ptoFromSheetName);
const ptoFormsSheetValues = ptoFormsSheet.getDataRange().getValues();

/**
 * Add one to days not approved.
 *
 */
// eslint-disable-next-line no-unused-vars, require-jsdoc
function addOneToDaysNotApproved() {
  for ( let i = 1; i < ptoFormsSheetValues.length; i++ ) {
    const currentValue =
      ptoFormsSheetValues[i][PTOFORMS_DAYSNOTAPPROVEDCOLUMN - 1];

    ptoFormsSheet.getRange(i + 1, PTOFORMS_DAYSNOTAPPROVEDCOLUMN)
        .setValue(currentValue + 1);

    if ( currentValue > 4 ) {
      lineManagerReminder(i);
    }
  }
}

/**
 * Send linemanager a reminder email.
 *
 * @param {number} i
 */
// eslint-disable-next-line no-unused-vars, require-jsdoc
function lineManagerReminder(i) {
  MailApp.sendEmail({
    to: ptoFormsSheetValues[i][PTOFORMS_LINEMANAGERCOLUMN - 1],
    subject: 'PTO Approval Reminder for ' +
      ptoFormsSheetValues[i][PTOFORMS_USERNAME - 1],
    htmlBody: 'A PTO request has been pending for 5 or more days.' +
      'The form can be found here:<br>' +
      ptoFormsSheetValues[i][PTOFORMS_APPROVALFORMCOLUMN - 1],
  });
}
