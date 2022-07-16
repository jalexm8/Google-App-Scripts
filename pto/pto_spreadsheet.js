const PTOFORMS_USERNAME              = 1;
const PTOFORMS_LINEMANAGERCOLUMN     = 3;
const PTOFORMS_APPROVALFORMCOLUMN    = 6;
const PTOFORMS_DAYSNOTAPPROVEDCOLUMN = 7;

var spreadSheetID         = "1HsG9B7Mrk_oJ6cLfaPoX9FyGwHZnp42Y_TGK-AoG9HU";
var ptoFromSheetName      = "pendingPtoApprovalForms";
var spreadSheet           = SpreadsheetApp.openById(spreadSheetID);
var ptoFormsSheet         = spreadSheet.getSheetByName(ptoFromSheetName);
var ptoFormsSheetValues   = ptoFormsSheet.getDataRange().getValues();

function addOneToDaysNotApproved() {
  for ( var i = 1; i < ptoFormsSheetValues.length ; i++ ) {
    var currentValue = ptoFormsSheetValues[i][PTOFORMS_DAYSNOTAPPROVEDCOLUMN - 1];

    ptoFormsSheet.getRange(i + 1, PTOFORMS_DAYSNOTAPPROVEDCOLUMN).setValue(currentValue + 1);

    if ( currentValue > 4 ) {
      lineManagerReminder(i);
    }
  }
}

function lineManagerReminder(i) {
  MailApp.sendEmail({
      to: ptoFormsSheetValues[i][PTOFORMS_LINEMANAGERCOLUMN - 1],
      subject: "PTO Approval Reminder for " + ptoFormsSheetValues[i][PTOFORMS_USERNAME - 1],
      htmlBody: "A PTO request has been pending for 5 or more days. The form can be found here:<br>" +
                ptoFormsSheetValues[i][PTOFORMS_APPROVALFORMCOLUMN - 1]
  });
}
