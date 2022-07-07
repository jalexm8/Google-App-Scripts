function formResponsesToArray() {
    var form = FormApp.getActiveForm();
    var formResponses = form.getResponses();
    var lastResponse = formResponses[formResponses.length - 1].getItemResponses();
    var userEmail = formResponses[formResponses.length - 1].getRespondentEmail();
    var startDate = lastResponse[0].getResponse();
    var endDate = lastResponse[1].getResponse();
  
    var userResponses = {
      'email': userEmail,
      'start date': startDate,
      'end date': endDate,
    };
  
    return userResponses;
  }
  
function onFormSubmit(e) {
    var userResponses = formResponsesToArray();
  
    Logger.log(userResponses['email'])
    Logger.log(userResponses['start date']);
    Logger.log(userResponses['end date']);
  }
  