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
  
  function errorChecking(userResponses) {
    var startDate = new Date(userResponses['start date']);
    var endDate = new Date(userResponses['end date']);
    var errorMsg = "";
  
    if ( startDate.getTime() > endDate.getTime() ) {
      errorMsg += "ERROR: Start date cannot be ahead of end date.\n";
    }
    
    Logger.log(errorMsg);
  
    return errorMsg;
  }
  
  function daysBetweenStartAndEndDate(userResponses) {
    var startDate = new Date(userResponses['start date']);
    var endDate = new Date(userResponses['end date']);
    var timeBetween = endDate.getTime() - startDate.getTime();
    var daysBetween = timeBetween / (1000 * 3600 * 24);
  
    return daysBetween;
  }
  
  function onFormSubmit(e) {
    var userResponses = formResponsesToArray();
    var ptoDaysTaken = daysBetweenStartAndEndDate(userResponses);
    var errorMsg = errorChecking(userResponses);
  
    Logger.log(userResponses['email'])
    Logger.log(userResponses['start date']);
    Logger.log(userResponses['end date']);
    Logger.log(ptoDaysTaken);
  
  
    if ( errorMsg.length > 0 ) {
      Logger.log(errorMsg);
    }
  }
  