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
function LogThrottled(id, logMessage, verbose, throttleOffset)
{
  const defaultThrottleOffset= 10 * 60;
  const throttledTime= GetValueByName(id, "ParameterAlertThrottleTime", verbose);
  const currentTime= new Date();
  
  if (verbose == undefined)
  {
    verbose = false;
  }

  if (throttleOffset == undefined)
  {
    throttleOffset= defaultThrottleOffset;
  }
  
  if (currentTime > throttledTime)
  {
    Logger.log(`[${LogThrottled.caller.name}] ${logMessage}`);
    ThrottleLog(id, throttleOffset, verbose);
  }
  else if (verbose)
  {
    Log(`Log messages throttled until ${throttledTime}`);
  }
};
  
  
/**
 * ThrottleLog()
 *
 * Set a time until which throttled log messages will be supporessed
 */
function ThrottleLog(id, untilTimeOffset, verbose)
{
  const defaultThrottleOffset= 10 * 60;
  const throttleTime= new Date();

  if (untilTimeOffset == undefined)
  {
    untilTimeOffset= defaultThrottleOffset;
  }

  if (verbose == undefined)
  {
    verbose = false;
  }

  throttleTime.setSeconds(throttleTime.getSeconds() + untilTimeOffset);

  SetValueByName(id, "ParameterAlertThrottleTime", throttleTime, verbose);
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