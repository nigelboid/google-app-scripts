/**
 * Log()
 *
 * Create a log entry with a label from the calling function
 */
function Log(logMessage)
{
  Logger.log(`[${Log.caller.name}] ${logMessage}`);
};
  
  
/**
 * LogVerbose()
 *
 * Create a log entry with a label from the calling function
 */
function LogVerbose(logMessage, verbose)
{
  if (verbose)
  {
    Logger.log(`[${LogVerbose.caller.name}] ${logMessage}`);
  }
};
  
  
/**
 * LogThrottled()
 *
 * Create a log entry with a label from the calling function if outside the throttled window
 */
function LogThrottled(sheetID, logMessage, verbose, throttleOffset)
{
  const defaultThrottleOffset = 10 * 60;
  const throttledTime = GetValueByName(sheetID, "ParameterAlertThrottleTime", verbose);
  const currentTime = new Date();
  
  if (verbose == undefined)
  {
    verbose = false;
  }

  if (throttleOffset == undefined)
  {
    throttleOffset = defaultThrottleOffset;
  }
  
  if (currentTime > throttledTime)
  {
    Logger.log(`[${LogThrottled.caller.name}] ${logMessage}`);
    SetValueByName(sheetID, "ParameterAlertThrottleMessage", logMessage, verbose);
    ThrottleLog(sheetID, throttleOffset, verbose);
  }
  else if (verbose)
  {
    Log(`Log messages throttled until ${throttledTime}`);
  }
};
  
  
/**
 * ThrottleLog()
 *
 * Set a time until which throttled log messages will be suppressed
 */
function ThrottleLog(sheetID, untilTimeOffset, verbose)
{
  const defaultThrottleOffset = 10 * 60;
  const throttleTime = new Date();

  if (untilTimeOffset == undefined)
  {
    untilTimeOffset = defaultThrottleOffset;
  }

  if (verbose == undefined)
  {
    verbose = false;
  }

  throttleTime.setSeconds(throttleTime.getSeconds() + untilTimeOffset);

  SetValueByName(sheetID, "ParameterAlertThrottleTime", throttleTime, verbose);
  Log(`Throttling further log messages until ${throttleTime}`)
};
  
  
/**
 * LogSend()
 *
 * Send the log as an e-mail message to the script invoker
 */
function LogSend(sheetID)
{
  // declare local variables and constants
  var recipient = null;
  var subject = 'Log';
  var body = Logger.getLog();
  var spreadsheet = null;
  var verbose = false;
 
  if (body)
  {
    // looks like we have something to send
    recipient = GetValueByName(sheetID, "ParameterAlertEmailAddress", verbose);
    if (!recipient)
    {
      // Unpsecified log recipient, alert the user
      Logger.log("[LogSend] No log recipient specified for spreadsheet ID <%s>; alerting the active user instead", sheetID);
      
      recipient = Session.getActiveUser().getEmail();
    }
    
    
    if (sheetID)
    {
      if (spreadsheet = SpreadsheetApp.openById(sheetID))
      {
        subject= spreadsheet.getName() + " " + subject;
      }
      else
      {
        Logger.log("[LogSend] Could not open spreadsheet ID <%s>.", sheetID);
        subject= sheetID + " " + subject;
      }
    }
    else
    {
      Logger.log("[LogSend] Did not receive a valid ID <%s>.", sheetID);
    }
    
    
    MailApp.sendEmail(recipient, subject, Logger.getLog());
    Logger.clear();
  }
};