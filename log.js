/**
 * Log()
 *
 * Create a log entry woth a label from the calling function
 */
function Log()
{
  Logger.log(`[${Log.caller.name}] ${arguments[0]}`);
};
  
  
/**
 * LogSend()
 *
 * Send the log as an e-mail message to the script invoker
 */
function LogSend(id)
{
  // declare local variables and constants
  var recipient= null;
  var subject= 'Log';
  var body= Logger.getLog();
  var spreadsheet= null;
  var verbose= false;
 
  if (body)
  {
    // looks like we have something to send
    recipient= GetValueByName(id, "ParameterAlertEmailAddress", verbose);
    if (!recipient)
    {
      // Unpsecified log recipient, alert the user
      Logger.log("[LogSend] No log recipient specified for spreadsheet ID <%s>; alerting the active user instead", id);
      
      recipient= Session.getActiveUser().getEmail();
    }
    
    
    if (id)
    {
      if (spreadsheet= SpreadsheetApp.openById(id))
      {
        subject= spreadsheet.getName() + " " + subject;
      }
      else
      {
        Logger.log("[LogSend] Could not open spreadsheet ID <%s>.", id);
        subject= id + " " + subject;
      }
    }
    else
    {
      Logger.log("[LogSend] Did not receive a valid ID <%s>.", id);
    }
    
    
    MailApp.sendEmail(recipient, subject, Logger.getLog());
    Logger.clear();
  }
};