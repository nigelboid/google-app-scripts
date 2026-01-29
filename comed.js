/**
 * Main entry point for the continuous check
 *
 * Checks for fresh data and saves it once it is available
 */
function RunComEdFrequently()
{
  // Declare constants and local variables
  const verbose = false;
  var parameters = GetParametersComed(verbose);
  const lastStamp = GetLatestTimeStamp(parameters);
  const minutesElapsed = ConvertMillisecondsToMinutes(parameters["scriptTime"] - lastStamp);
  var success = false;

  // Since ComEd provides 5-minutes prices and provides them about 8 minutes late,
  //   only run if more than 8 minutes have elapsed since the last entry and if not already running
  if (lastStamp && parameters["scriptTime"] > (lastStamp + parameters["checkThreshold"]))
  {
    UpdateStatus(parameters, "Getting prices...");
    const prices = GetPricesComed(parameters, lastStamp);
    
    if (prices.length > 0)
    {
      // Looks like we have data, proceed
      UpdateStatus(parameters, "Getting prices: Latest prices obtained!");
      
      if (success = PrepareToCommit(parameters))
      {
        if (success = SavePrices(parameters, prices))
        {
          if (parameters["noNewPricesAlert"])
          {
            Log(`New prices updated after a delay (${minutesElapsed} minutes, ${parameters["priceLast"]}¢)`);
            // LogVerbose(`New prices updated after a delay (${minutesElapsed} minutes, ${parameters["priceLast"]}¢)`, verbose);
            ClearMissingPricesAlertStamp(parameters);
          }
            
          if (success = UpdateComputedValues(parameters))
          {
            success = Notify(parameters);
          }
        }
      }
      
      // Clean up
      if (success)
      {
        // Complete and successful new price commitment: attempt to clear semaphore
        success = ClearSemaphore(parameters);
        if(!success)
        {
          Log(`Failed to clear current semaphore (${parameters["scriptTime"].toFixed(0)}) upon completion!`);
          // LogVerbose(`Failed to clear current semaphore (${parameters["scriptTime"].toFixed(0)}) upon completion!`, verbose);
        }
      }
      else
      {
        // Report early and abnormal termination
        success = Scuttle(parameters, "RunComEdFrequently");
      }
    }
    else
    {
      const action = "No new prices obtained";
      
      if (parameters["scriptTime"] > (lastStamp + parameters["checkThreshold"] * 2))
      {
        if (parameters["noNewPricesAlert"].length == 0)
        {
          Log(`No new prices available (${minutesElapsed}`);
          // LogVerbose(`No new prices available (${minutesElapsed}`, verbose);
          UpdateMissingPricesAlertStamp(parameters);
        }
        
        UpdateStatus(parameters, action + ": Stale prices!");
      }
      else
      {
        UpdateStatus(parameters, action + ": Scrubbing history while waiting for updated prices...");
        ScrubHistory(parameters, action);
      }
    }
  }
  else
  {
    var action = "Invoked too soon";
    
    UpdateStatus(parameters, action + ": Scrubbing history instead of checking prices...");
    ScrubHistory(parameters, action);
  }
  
  LogSend(parameters["sheetID"]);
  return success;
};


/**
 * GetParametersComed()
 *
 * Create and return a hash of various parameters
 */
function GetParametersComed(verbose)
{
  const id = "1ZodNejQNITpXlsj4ccN7h_ttHKxP4MviYD-kywFKbOY";
  var parameters = GetParameters(id, "comedParameters", verbose);

  parameters["scriptTime"] = new Date().getTime();
  
  parameters["confirmNumbers"] = true;
  parameters["confirmNumbersLimit"] = -1000000;
  
  parameters["indexTime"] = 0;
  parameters["indexPrice"] = parameters["indexTime"] + 1;
  parameters["indexMovingAverage"] = parameters["indexPrice"] + 1;
  parameters["indexTrend"] = parameters["indexMovingAverage"] + 1;
  parameters["indexStamp"] = parameters["indexTrend"] + 1;
  parameters["indexStampTime"] = parameters["indexStamp"] + 1;
  parameters["indexAlert"] = parameters["indexStampTime"] + 1;
  parameters["priceTableWidth"] = parameters["indexAlert"] + 1;
  
  LogVerbose(`Parameters:\n\n${parameters}`, verbose);
  
  return parameters;
};



/**
 * GetPricesComed()
 *
 * Grab pricing data from ComEd for the specified time interval
 */
function GetPricesComed(parameters, intervalStart)
{
  const timeKey = parameters["comedKeyTime"];
  const priceKey = parameters["comedKeyPrice"];
  const urlHead = parameters["comedURLHead"];
  var urlDateStart = parameters["comedURLDateRangeStart"];
  var urlDateEnd = parameters["comedURLDateRangeEnd"];
  const intervalEnd = new Date(parameters["scriptTime"]);
  
  // Formulate starting and ending date stamps, offset latest stamp by one minute
  intervalStart = new Date(intervalStart + 60000);
  
  urlDateStart += NumberToString(intervalStart.getFullYear(), 4, "0")
                + NumberToString(intervalStart.getMonth()+1, 2, "0")
                + NumberToString(intervalStart.getDate(), 2, "0")
                + NumberToString(intervalStart.getHours(), 2, "0")
                + NumberToString(intervalStart.getMinutes(), 2, "0");
  
  urlDateEnd += NumberToString(intervalEnd.getFullYear(), 4, "0")
              + NumberToString(intervalEnd.getMonth() + 1, 2, "0")
              + NumberToString(intervalEnd.getDate(), 2, "0")
              + NumberToString(intervalEnd.getHours(), 2, "0")
              + NumberToString(intervalEnd.getMinutes(), 2, "0");
  
  
  // Obtain and parse missing pricing data
  const options = {'muteHttpExceptions' : true };
  const responseOk = 200;
  const url = urlHead + urlDateStart + urlDateEnd;
  
  UpdateURL(parameters, url);
  
  const response = UrlFetchApp.fetch(url, options);
  const responseCode = response.getResponseCode();
  var priceTable= [];
  
  if (responseCode == responseOk)
  {
    // looks like we received a benign response
    var prices = JSON.parse(response.getContentText());
    var row = null;
    var timeStamp = null;
    
    prices.sort(function(a, b){return a[timeKey] - b[timeKey]});
    for (const price of prices)
    {
      // convert each data line into a padded table row
      row = FillArray(parameters["priceTableWidth"], "");
      timeStamp = new Date();
      timeStamp.setTime(price[timeKey]);
      row[parameters["indexTime"]] = timeStamp;
      row[parameters["indexPrice"]] = price[priceKey];
      row = priceTable.push(row);
    }
    
    if (row > 0)
    {
      // Update status with latest price and time
      UpdateLastPrice(parameters, prices[row-1][priceKey]);
    }
  }
  else
  {
    // looks like we did not obtain our prices
    LogVerbose(`Asked for latest prices, but received an unexpected response code instead: <${responseCode.toFixed(0)}`,
                true
                // parameters["verbose"]
              );
  }
  
  return priceTable;
};


/**
 * SendPriceAlert()
 *
 * Send a price alert, depending on conditions
 */
function SendPriceAlert(parameters)
{
  const priceCurrent = parameters["priceLast"];
  const priceMovingAverage = parameters["priceMovingAverage"];
  const priceTrend = parameters["priceRegressionSlope"];;
  
  const priceExpensive = parameters["priceLimitExpensive"];
  const priceNormalUpper = parameters["priceLimitNormalUpper"];
  const priceNormalLower = parameters["priceLimitNormalLower"];
  const priceCheap = parameters["priceLimitCheap"];
  const priceThresholdMovingAverage = parameters["priceThresholdMovingAverage"];
  const priceThresholdStable = parameters["priceThresholdStable"];
  const priceAlertLast = parameters["priceAlertLast"];
  const priceThresholdAlert = parameters["priceThresholdAlert"];
  const priceAlertDNDStart = parameters["priceAlertDNDStart"];
  const priceAlertDNDEnd = parameters["priceAlertDNDEnd"];
  const priceAlertDNDTemp = parameters["priceAlertDNDTemp"];
 
  var alert= priceCurrent + "¢ (Δ= " + (priceCurrent-priceAlertLast).toFixed(1) + "¢ [±" + priceThresholdAlert + "¢]) and ";
  var status= "No one should ever see this!";
  
  UpdateStatus(parameters, "Composing trend missive...");
  
  if ((priceCurrent - priceMovingAverage) > priceThresholdStable)
  {
    alert+= "rising";
  }
  else if ((priceCurrent - priceMovingAverage) < -priceThresholdStable)
  {
    alert+= "falling";
  }
  else
  {
    alert+= "steady";
  }
  alert += " (inflection= " + (priceTrend * parameters["priceRegressionSlope"]).toFixed(0);
  alert += ", MA deviation= " + (priceCurrent - priceMovingAverage).toFixed(2) + "¢ [±" + priceThresholdMovingAverage + "¢]" + ")";
  
  if ((Math.abs(priceCurrent-priceAlertLast) > priceThresholdAlert)
      && ((priceTrend * parameters["priceRegressionSlope"] <= 0) || (Math.abs(priceCurrent - priceMovingAverage) > priceThresholdMovingAverage)))
  {
    // Price jumping too much since last alert and either inflecting trend (slope) or deviating too far from the moving average -- check if that merits an alert...
    // ...may trigger on a flat slope -- that's a feature!
    UpdateStatus(parameters, "Composing alert...");
    
    if (priceCurrent > priceExpensive)
    {
      // High price range
      alert = "Expensive electricity: " + alert;
    }
    else if (priceCurrent < priceCheap)
    {
      // Low price range
      alert = "Cheap electricity: " + alert;
    }
    else
    {
      // Situation within or near normal bounds -- suppress alert???
      if (priceCurrent < priceNormalLower)
      {
        // Just below normal
        alert = "Now just below normal: " + alert;
      }
      else if (priceCurrent > priceNormalUpper)
      {
        // Just above normal
        alert = "Now just above normal: " + alert;
      }
      else
      {
        // Normal
        alert = "Now normal: " + alert;
      }
    }
  }
  else
  {
    // Prices not moving much within current bounds -- suppress alert
    alert= "No alert triggered (parameters stable): " + alert;
    if (!parameters["verbose"])
    {
      status = alert;
      alert = null;
    }
  }
  
  if (alert)
  {
    // Alert condition reached -- apply cosmetics and check suppressed alert window
    const alertTime = new Date();
    const hour = alertTime.getHours();
    
    UpdateStatus(parameters, "Alert composed.");
    UpdateAlertStatus(parameters, priceCurrent, alert, alertTime);
    status = alert;
    
    if (priceAlertDNDStart > priceAlertDNDEnd)
    {
      // Adjust for crossing the day boundary
      priceAlertDNDEnd += 24;
      if (hour < priceAlertDNDStart)
      {
        hour += 24;
      }
    }
    
    if (((hour >= priceAlertDNDStart) && (hour < priceAlertDNDEnd)) || (priceAlertDNDTemp > alertTime))
    {
      // Suppress an actual alert during the Do Not Disturb Window, but annotate the status
      status += " [suppressed]";
    }
    else
    {
      // Trigger an actual alert since we are outside the Do Not Disturb window
      Log(`${alert}`);
      LogVerbose(
                  `Current hour [${hour}] is outside the do not disturb window [${priceAlertDNDStart} - ${priceAlertDNDEnd}].`,
                  true
                  // parameters["verbose"]
                );
    }
  }
  
  return status;
};


/**
 * GetLatestTimeStamp()
 *
 * Get the latest time stamp from the history table
 */
function GetLatestTimeStamp(parameters)
{
  const stamp = GetLastSnapshotStamp(parameters["sheetID"], parameters["comedSheetPrices"], parameters["verbose"]);
  
  if (stamp && (stamp.toString().length > 0))
  {
    return new Date(stamp).getTime();
  }
  else
  {
    Log(`Retrieved an invalid stamp [${stamp}].`);
    // LogVerbose(`Retrieved an invalid stamp [${stamp}].`, parameters["verbose"]);
    SetLatestTimeStamp(parameters);
  }
  
  return null;
};


/**
 * SetLatestTimeStamp()
 *
 * Set the latest time stamp in the history table (error condition recovery)
 */
function SetLatestTimeStamp(parameters)
{
  var stamp = null;
  const onlyIfBlank = false;
  
  if (stamp = new Date())
  {
    const success = UpdateSnapshotCell(
                                        parameters["sheetID"],
                                        parameters["comedSheetPrices"],
                                        parameters["indexTime"] + 1,
                                        stamp,
                                        onlyIfBlank,
                                        parameters["verbose"]
                                      );
    if (success)
    {
      Log(`Overwrote latest time stamp with [${stamp}].`);
    }
  }
  else
  {
    Log(`Could not overwrite latest time stamp.`);
  }
};


/**
 * ScrubHistory()
 *
 * Remove duplicate rows from history (due to superseded runs?) and otherwise keep the history to a maximum number of entries
 */
function ScrubHistory(parameters, action)
{
  var scrubbedData = null;
  var semaphore = null;
  const maxRows = 3000;
  
  if (semaphore = GetSemaphore(parameters))
  {
    // Semaphore precludes scrubbing -- preserve its value for the chain of command and restore prior status
    const statusAction = "Deferred";
    const statusDetails = SemaphoreConflictDetails(parameters, semaphore);
      
    PreserveStatus(parameters, statusAction);
    Log(`Deferring scrubbing history (${statusDetails})...`);
    // LogVerbose(`Deferring scrubbing history (${statusDetails})...`, parameters["verbose"]);
  }
  else
  {
    // Check for a duplicate snashot row and preserve its values for the chain of command
    if (scrubbedData = RemoveDuplicateSnapshot(parameters["sheetID"], parameters["comedSheetPrices"], parameters["verbose"]))
    {
      UpdateStatus(parameters, "Removed duplicate history row.");
      Log(`Removed duplicate history row\n\n${scrubbedData}`);
      // LogVerbose(`Removed duplicate history row\n\n${scrubbedData}`, parameters["verbose"]);
    }
    else
    {
      TrimHistory(parameters["sheetID"], parameters["comedSheetPrices"], maxRows, parameters["verbose"]);
      UpdateStatus(parameters, action + ".");
    }
  }
  
  return scrubbedData;
};


/**
 * IsSupreme()
 *
 * Determine if another run has superseded this one (via run time stamps)
 */
function IsSupreme(parameters)
{
  const current = GetValueByName(parameters["sheetID"], "statusRunCurrent", parameters["verbose"]);
  
  return (parameters["scriptTime"] >= current);
}


/**
 * Superseded()
 *
 * Report that this run has been superseded by another
 */
function Superseded(parameters, caller, activity)
{
  const current = GetValueByName(parameters["sheetID"], "statusRunCurrent", parameters["verbose"]);
  var statusMessage = "Superseded!";
  var logMessage = "";
  
  // Report a stale run
  logMessage = "Superseded by a later run ("
              + current.toFixed(0)
              + " started "
              + ConvertMillisecondsToMinutes(current - parameters["scriptTime"])
              + " minutes later; current= "
              + parameters["scriptTime"].toFixed(0)
              + ")";
  
  if (activity)
  {
    // Insert status information into status and log missives
    statusMessage = activity + ": " + statusMessage;
    logMessage += " while " + activity.toLowerCase() + ".";
  }
  
  if (caller)
  {
    // Prepend caller indentifier to log missive
    logMessage = "[" + caller + "] " + logMessage;
  }
  
  UpdateStatus(parameters, statusMessage);
  Log(`${logMessage}`);
  // LogVerbose(`${logMessage}`, parameters["verbose"]);
};


/**
 * GetSemaphore()
 *
 * Obtain the latest semaphore, if any
 */
function GetSemaphore(parameters)
{
  return GetValueByName(parameters["sheetID"], "semaphore", parameters["verbose"]);
};


/**
 * SetSemaphore()
 *
 * Set a semaphore since this run is writing to history
 */
function SetSemaphore(parameters)
{
  var semaphore = null;
  
  semaphore = GetSemaphore(parameters);
  if (semaphore)
  {
    // Blocked by another run
    const semaphoreDetails = SemaphoreConflictDetails(parameters, semaphore);

    Log(`Could not set semaphore (${semaphoreDetails})!`);
    // LogVerbose(`Could not set semaphore (${semaphoreDetails})!`, parameters["verbose"]);
    Log(`Prior status: ${parameters["status"]}`);
    // LogVerbose(`Prior status: ${parameters["status"]}`, parameters["verbose"]);
    
    if (parameters["forceThreshold"] && (parameters["scriptTime"] - semaphore) > parameters["forceThreshold"])
    {
      // Enough time has elapsed for us to forcefully clear a prior semaphore
      var force = true;
      
      UpdateStatus(parameters, "Forcefully clearing prior semaphore...");
      Log(`Attempting to forcefully clear prior semaphore (${semaphore.toFixed(0)})...`);
      // LogVerbose(`Attempting to forcefully clear prior semaphore (${semaphore.toFixed(0)})...`, parameters["verbose"]);
      
      force = ClearSemaphore(parameters, force);
      if (force)
      {
        const action = "Forcefully cleared prior semaphore";
        
        UpdateStatus(parameters, action + ": Scrubbing history...");
        // Stale semaphore cleared -- check for cobbled data
        ScrubHistory(parameters, action)
      }
      else
      {
        Log(`Failed to forcefully clear prior semaphore!`);
        // LogVerbose(`Failed to forcefully clear prior semaphore!`, parameters["verbose"]);
      }
    }
    
    return false;
  }
  else
  {
    // Clear to proceed
    SetValueByName(parameters["sheetID"], "semaphore", parameters["scriptTime"].toFixed(0), parameters["verbose"]);
    SetValueByName(parameters["sheetID"], "semaphoreTime", DateToLocaleString(), parameters["verbose"]);
    parameters["semaphore"] = parameters["scriptTime"];
    
    return true;
  }
};


/**
 * ClearSemaphore()
 *
 * Clear our semaphore (writing)
 */
function ClearSemaphore(parameters, force)
{
  const semaphore = GetSemaphore(parameters);
  var success = false;
  
  if (semaphore)
  {
    // There is a semaphore set -- confirm it is current and proceed accordingly
    if (IsSupreme(parameters))
    {
      // Proceed to clear semaphore
      if (parameters["scriptTime"] == semaphore)
      {
        // Normally only clear own semaphore
        success = SetValueByName(parameters["sheetID"], "semaphore", "", parameters["verbose"]);
        SetValueByName(parameters["sheetID"], "semaphoreTime", DateToLocaleString(), parameters["verbose"]);
      }
      else if (force)
      {
        // Clear another run's semaphore if set to do so
        success = SetValueByName(parameters["sheetID"], "semaphore", "", parameters["verbose"]);
        SetValueByName(parameters["sheetID"], "semaphoreTime", DateToLocaleString(), parameters["verbose"]);
        Log(`Clearing a semaphore from another run (${semaphore.toFixed(0)})!`);
        // LogVerbose(`Clearing a semaphore from another run (${semaphore.toFixed(0)})!`, parameters["verbose"]);
      }
      else
      {
        const semaphoreDetails = SemaphoreConflictDetails(parameters, semaphore);
        success = false;
        Log(`Cannot clear a semaphore from another run (${semaphoreDetails})!`);
        // LogVerbose(`Cannot clear a semaphore from another run (${semaphoreDetails})!`, parameters["verbose"]);
      }
    }
    else
    {
      // Technically, should never enter this clause due to semaphore and precedence (superseded) checks -- or so I had thought!
      const semaphoreDetails = SemaphoreConflictDetails(parameters, semaphore);
      success = false;
      Log(`Cannot clear a semaphore since another run superseded this one (${semaphoreDetails})!`);
      // LogVerbose(`Cannot clear a semaphore since another run superseded this one (${semaphoreDetails})!`, parameters["verbose"]);
    }
  }
  else
  {
    // No semaphore set?!
    success = false;
    LogVerbose(
                `Something or someone else has already cleared the semaphore (${parameters["scriptTime"].toFixed(0)})!`,
                true
                // parameters["verbose"]
              );
  }
  
  return success;
};


/**
 * Scuttle()
 *
 * Deal with an abnormal and early termination
 */
function Scuttle(parameters, caller)
{
  var success = false;
  var logMessage = "Scuttling";
  const statusAction = "Scuttled";
  
  UpdateStatus(parameters, "Scuttling...");
  
  // Compose and commit log message
  if (caller != undefined)
  {
    logMessage= "[" + caller + "] " + logMessage;
  }
  
  if (parameters["activity"])
  {
    logMessage+= " while " + parameters["activity"] + ".";
  }
  else
  {
    logMessage+= ".";
  }
  
  Log(logMessage);
  // LogVerbose(logMessage, parameters["verbose"]);
  
  // Clean up
  if (parameters["semaphore"] == parameters["scriptTime"])
  {
    success= ClearSemaphore(parameters);
    if (!success)
    {
      Log(`Failed to clear current semaphore (${parameters["scriptTime"].toFixed(0)})!`);
      // LogVerbose(`Failed to clear current semaphore (${parameters["scriptTime"].toFixed(0)})!`, parameters["verbose"]);
    }
    else
    {
      // Scuttling is a failure regardless of intermediate successes!
      success = false;
    }
  }
  
  // Restore prior status for consistency and report failure
  PreserveStatus(parameters, statusAction);
  return success;
};


/**
 * SemaphoreConflictDetails()
 *
 * Compose string descricing semaphore conflict details
 */
function SemaphoreConflictDetails(parameters, semaphore)
{
  const deltaValue = ConvertMillisecondsToMinutes(parameters["scriptTime"] - semaphore);
  var details = "blocked by semaphore ";
  
  details = details.concat(semaphore.toFixed(0), " set ");
  
  if (deltaValue < 0)
  {
    // Conflicting semaphore set later
    details = details.concat(-deltaValue, " minutes later");
  }
  else
  {
    // Conflicting semaphore set earlier or concurrently
    details = details.concat(deltaValue, " minutes earlier");
  }
  
  details = details.concat("; current: ", parameters["scriptTime"].toFixed(0));
  
  return details;
};


/**
 * UpdateStatus()
 *
 * Update status and time
 */
function UpdateStatus(parameters, status)
{
  SetValueByName(parameters["sheetID"], "status", status, parameters["verbose"]);
  SetValueByName(parameters["sheetID"], "statusTime", DateToLocaleString(), parameters["verbose"]);
};


/**
 * PreserveStatus()
 *
 * Preserve previous status
 */
function PreserveStatus(parameters, statusAction)
{
  var statusPrior = parameters["status"];
  const statusKeys = ["Deferred", "Scuttled"];
  const statusPreambleFiller = " due to: ";
  const statusPreamble = statusAction + statusPreambleFiller;
  
  if (!statusPrior.includes(statusPreamble))
  {
    // Add preamble since it is not included yet
    statusPrior = statusPreamble + statusPrior;
  }
  else
  {
    // Update an existing preamble
    for (const statusKey of statusKeys)
    {
      // Check if another action preserved status prior to this attempt
      if (statusPrior.includes(statusKey + statusPreambleFiller))
      {
        if (statusKey != statusAction)
        {
          // Replace the previous, non-matching  preserving action with the current one
          statusPrior.replace(statusKey + statusPreambleFiller, statusPreamble);
        }
        
        break;
      }
    }
  }
  
  UpdateStatus(parameters, statusPrior);
};


/**
 * UpdateAlertStatus()
 *
 * Update status and time of alert
 */
function UpdateAlertStatus(parameters, price, alert, time)
{
  SetValueByName(parameters["sheetID"], "statusAlert", alert, parameters["verbose"]);
  SetValueByName(parameters["sheetID"], "statusAlertTime", DateToLocaleString(time), parameters["verbose"]);
  SetValueByName(parameters["sheetID"], "priceAlertLast", price, parameters["verbose"]);
};


/**
 * UpdateRunStamps()
 *
 * Update status and time of status
 */
function UpdateRunStamps(parameters)
{
  const success = SetValueByName(
                                  parameters["sheetID"],
                                  "statusRunCurrent",
                                  parameters["scriptTime"].toFixed(0),
                                  parameters["verbose"]
                                );
  
   if (success)
   {
     SetValueByName(
                    parameters["sheetID"],
                    "statusRunCurrentTime",
                    DateToLocaleString(parameters["scriptTime"]),
                    parameters["verbose"]
                  );
     SetValueByName(
                    parameters["sheetID"],
                    "statusRunPrevious",
                    parameters["statusRunCurrent"],
                    parameters["verbose"]
                  );
     SetValueByName(
                    parameters["sheetID"],
                    "statusRunPreviousTime",
                    DateToLocaleString(parameters["statusRunCurrent"]),
                    parameters["verbose"]
                  );
   }
  
  return success;
};


/**
 * UpdatePreviousRegressionSlope()
 *
 * Preserve previous regression slope
 */
function UpdatePreviousRegressionSlope(parameters)
{
  UpdateStatus(parameters, "Saving previous regression slope...");
  
  const success = SetValueByName(
                                  parameters["sheetID"],
                                  "priceRegressionSlopePrevious",
                                  parameters["priceRegressionSlope"],
                                  parameters["verbose"]
                                );
  
  if (success)
  {
    SetValueByName(parameters["sheetID"], "priceRegressionSlopePreviousTime", DateToLocaleString(), parameters["verbose"]);
  }
  
  return success;
};


/**
 * UpdateLastPrice()
 *
 * Preserve latest price
 */
function UpdateLastPrice(parameters, price)
{
  const success = SetValueByName(parameters["sheetID"], "priceLast", price, parameters["verbose"]);
  
  if (success)
  {
    parameters["priceLast"] = price;
    SetValueByName(parameters["sheetID"], "priceLastTime", DateToLocaleString(), parameters["verbose"]);
  }
  
  return success;
};


/**
 * UpdateURL()
 *
 * Preserve latest query URL
 */
function UpdateURL(parameters, url)
{
  SetValueByName(parameters["sheetID"], "comedURL", url, parameters["verbose"]);
  SetValueByName(parameters["sheetID"], "comedURLTime", DateToLocaleString(), parameters["verbose"]);
};


/**
 * UpdateMissingPricesAlertStamp()
 *
 * Preserce the step of the last time we alerted about missing (delayed) prices
 */
function UpdateMissingPricesAlertStamp(parameters)
{
  SetValueByName(parameters["sheetID"], "noNewPricesAlert", parameters["scriptTime"], parameters["verbose"]);
  SetValueByName(parameters["sheetID"], "noNewPricesAlertTime", DateToLocaleString(), parameters["verbose"]);
};


/**
 * ClearMissingPricesAlertStamp()
 *
 * Preserce the step of the last time we alerted about missing (delayed) prices
 */
function ClearMissingPricesAlertStamp(parameters)
{
  SetValueByName(parameters["sheetID"], "noNewPricesAlert", "", parameters["verbose"]);
  SetValueByName(parameters["sheetID"], "noNewPricesAlertTime", DateToLocaleString(), parameters["verbose"]);
};


/**
 * PrepareToCommit()
 *
 * Prepare for writing and alerting
 */
function PrepareToCommit(parameters)
{
  var success = false;
  var activity = "Preparing to commit";
  
  UpdateStatus(parameters, activity + "...");
  parameters["activity"] = activity.toLowerCase();
  
  if (success = IsSupreme(parameters))
  {
    // This is still the latest run: grab and preserve latest regression slope and moving average values
    if (success = SetSemaphore(parameters))
    {
      // No conflicting runs -- proceed
      if (success = UpdateRunStamps(parameters))
      {
        success = UpdatePreviousRegressionSlope(parameters);
      }
    }
  }
  else
  {
    Superseded(parameters, "PrepareToCommit", activity);
  }
  
  return success;
};


/**
 * SavePrices()
 *
 * Save obtained prices
 */
function SavePrices(parameters, prices)
{
  var success = false;
  const updateRun = false;
  const activity = "Saving prices";
  
  UpdateStatus(parameters, activity + "...");
  parameters["activity"] = activity.toLowerCase();
  
  if (success = IsSupreme(parameters))
  {
    // This is still the latest run: write obtained prices to history table
    prices[prices.length-1][parameters["indexStamp"]] = parameters["scriptTime"].toFixed(0);
    prices[prices.length-1][parameters["indexStampTime"]] = DateToLocaleString();
    success = SaveSnapshot(parameters["sheetID"], parameters["comedSheetPrices"], prices, updateRun, parameters["verbose"]);
  }
  else
  {
    Superseded(parameters, "SavePrices", activity);
  }
  
  return success;
};


/**
 * UpdateComputedValues()
 *
 * Update trend and moving average
 */
function UpdateComputedValues(parameters)
{
  var success = false;
  var priceMovingAverage = null;
  var priceTrend = null;
  const onlyIfBlank = true;
  const activity = "Updating computed values";
  
  UpdateStatus(parameters, activity + "...");
  parameters["activity"] = activity.toLowerCase();
  
  if (success = IsSupreme(parameters))
  {
    // This is still the latest run: update freshly recomputed values
    parameters["activity"] = "updating moving average";
    
    priceMovingAverage = GetValueByName(parameters["sheetID"], "priceMovingAverage", parameters["verbose"], parameters["confirmNumbers"], parameters["confirmNumbersLimit"]);
    if (priceMovingAverage != null)
    {
      parameters["priceMovingAverage"] = priceMovingAverage;
      success = UpdateSnapshotCell(
                                    parameters["sheetID"],
                                    parameters["comedSheetPrices"],
                                    parameters["indexMovingAverage"] + 1,
                                    priceMovingAverage,
                                    onlyIfBlank,
                                    parameters["verbose"]
                                  );
    }
    else
    {
      success = false;
      Log(`Could not obtain updated Moving Average!"`);
      // LogVerbose(`Could not obtain updated Moving Average!"`, parameters["verbose"]);
    }
    
    if (success)
    {
      // Success so far: proceed...
      parameters["activity"] = "updating trend";
      
      priceTrend = GetValueByName(
                                  parameters["sheetID"],
                                  "priceRegressionSlope",
                                  parameters["verbose"],
                                  parameters["confirmNumbers"],
                                  parameters["confirmNumbersLimit"]
                                );

      if (priceTrend != null)
      {
        parameters["priceRegressionSlope"] = priceTrend;
        success = UpdateSnapshotCell(
                                      parameters["sheetID"],
                                      parameters["comedSheetPrices"],
                                      parameters["indexTrend"] + 1,
                                      priceTrend,
                                      onlyIfBlank,
                                      parameters["verbose"]
                                    );
      }
      else
      {
        success = false;
        Log(`Could not obtain updated Regression Coefficient!"`);
        // LogVerbose(`Could not obtain updated Regression Coefficient!"`, parameters["verbose"]);
      }
    }
  }
  else
  {
    Superseded(parameters, "SavePrices", activity);
  }
  
  return success;
};


/**
 * Notify()
 *
 * Notify via alerts, if triggered by parameters
 */
function Notify(parameters)
{
  var success = false;
  const onlyIfBlank = true;
  var statusMessage = null;
  var activity = "Preparing to notify";
  
  UpdateStatus(parameters, activity + "...");
  parameters["activity"] = activity.toLowerCase();
  
  if (success = IsSupreme(parameters))
  {
    // This is still the latest run: notify, if necessary
    statusMessage = SendPriceAlert(parameters);
    if (statusMessage.length > 0)
    {
      UpdateStatus(parameters, statusMessage);
      success = UpdateSnapshotCell(
                                    parameters["sheetID"],
                                    parameters["comedSheetPrices"],
                                    parameters["indexAlert"] + 1,
                                    statusMessage,
                                    onlyIfBlank,
                                    parameters["verbose"]
                                  );
    }
    else
    {
      Log(`Received empty status report <${statusMessage}> after trying to send price alert.`);
      // LogVerbose(`Received empty status report <${statusMessage}> after trying to send price alert.`, parameters["verbose"]);
    }
  }
  else
  {
    Superseded(parameters, "Notify", activity);
  }
  
  return success;
};


/**
 * ConvertMillisecondsToMinutes()
 *
 * Return elapsed time in minutes, converted from milliseconds
 */
function ConvertMillisecondsToMinutes(milliseconds)
{
  return (milliseconds / 60 / 1000).toFixed(2);
};