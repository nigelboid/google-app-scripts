/**
 * RunIndexStranglesCandidates()
 *
 * Find suitable contracts based on a list of parameters
 *
 */
function RunIndexStranglesCandidates(afterHours, test)
{
  // Declare constants and local variables
  const sheetID = GetMainSheetID();
  const forceRefreshNowName = "IndexStranglesForceRefreshNow";
  const optionPrices = true;
  const additionalCandidatesDTEDefault = 60;
  var verbose = false;
  var success = false;
  
  var forceRefreshNow = GetValueByName(sheetID, forceRefreshNowName, verbose);
  
  const nextUpdateTime = GetValueByName(sheetID, "IndexStranglesCandidatesUpdateNext", verbose);
  const currentTime = new Date();
  
  if (afterHours == undefined)
  {
    afterHours = true;
  }
  
  if (test == undefined)
  {
    test = false;
  }
  else
  {
    // Adjust other flags to trigger properly for testing
    verbose = test;
    forceRefreshNow = (forceRefreshNow || test);
  }

  if (forceRefreshNow)
  {
    // User set a manual forced refresh flag for index strangles...
    LogVerbose("Forcing a manual refresh of index strangles...", verbose);
    LogVerbose("Clearing the flag for manual refresh of index strangles...", verbose);
    
    SetValueByName(sheetID, forceRefreshNowName, "", verbose);
  }
  else
  {
    // Check for day changes
    if (currentTime.getDate() > nextUpdateTime.getDate())
    {
      // Force refresh on day changes
      forceRefreshNow = true;
    }
    else if (currentTime.getMonth() != nextUpdateTime.getMonth() && currentTime.getDate() == 1)
    {
      // Force refresh on day changes (special case for first day of the month)
      forceRefreshNow = true;
    }
  }
  
  if (forceRefreshNow || ((IsMarketOpen(sheetID, optionPrices, verbose) || afterHours) && (currentTime > nextUpdateTime)))
  {
    // Only check during market hours or after a day change
    var candidates = null;
    var candidatesAdditional = null;
    var dte = GetValueByName(sheetID, "IndexStranglesDTE", verbose);
    const symbols = GetTableByNameSimple(sheetID, "IndexStranglesSymbols", verbose);
    const deltaTargetCall = GetValueByName(sheetID, "IndexStranglesDeltaCall", verbose);
    const deltaTargetPut = GetValueByName(sheetID, "IndexStranglesDeltaPut", verbose);

    candidates = GetIndexStrangleContracts(sheetID, symbols, dte, deltaTargetCall, deltaTargetPut, verbose);

    //candidatesAdditional = GetIndexStrangleContractsHedge(sheetID, symbols, verbose);
    candidatesAdditional = GetIndexStrangleContracts(sheetID, symbols, dte, deltaTargetCall / 2, deltaTargetPut / 3, verbose);
    
    if (candidatesAdditional.length > 0)
    {
      // Add a blank line
      candidates.push([""]);
    
      // Add near-term candidates to the list of our primary candidates
      candidates = candidates.concat(candidatesAdditional);
    }
    else
    {
      Log(`Found no additional candidates <${candidatesAdditional}>!`);
    }

    if (candidates.length > 0)
    {
      // Looks like we obtained candidate contracts -- commit them to a table
      if (SetTableByName(sheetID, "IndexStranglesCandidates", candidates, verbose))
      {
        UpdateTime(sheetID, "IndexStranglesCandidatesUpdateTime", verbose);
        SetValueByName(sheetID, "IndexStranglesCandidatesUpdateStatus", "Updated [" + DateToLocaleString(currentTime) + "]", verbose);

        success = true;
      }
      else
      {
        SetValueByName(sheetID, "IndexStranglesCandidatesUpdateStatus",
                        "Failed to set new candidates [" + DateToLocaleString(currentTime) + "]", verbose);

        Log(`Failed to update table named 'IndexStranglesCandidates' in sheet <${sheetID}> with candidates: <${candidates}>!`);
      }
    }
    else
    {
      // No candidates found?
      UpdateTime(sheetID, "IndexStranglesCandidatesUpdateTime", verbose);
      SetValueByName(sheetID, "IndexStranglesCandidatesUpdateStatus",
                      "Failed to find new candidates [" + DateToLocaleString(currentTime) + "]", verbose);

      Log(`Found no candidates <${candidates}>!`);
      success = true;
    }
  }
  
  LogSend(sheetID);
  return success;
};


/**
 * GetIndexStrangleContracts()
 *
 * Obtain and select best matching strangles
 *
 */
function GetIndexStrangleContracts(sheetID, symbols, dte, deltaCall, deltaPut, verbose)
{
  var candidates= null;
  
  if (symbols && dte && deltaCall && deltaPut)
  {
    candidates= GetIndexStrangleContractsSchwabByDelta(sheetID, symbols, dte, deltaCall, deltaPut, verbose);
  }
  else
  {
    // Missing parameters
    candidates= [];
    Log(`Missing parameters: symbols= ${symbols}, DTE= ${dte}, calls delta= ${deltaCall}, puts delta= ${deltaPut}`);
  }

  return candidates;
};


/**
 * GetIndexStrangleContractsHedge()
 *
 * Obtain and select best matching strangles
 *
 */
function GetIndexStrangleContractsHedge(sheetID, symbols, verbose)
{
  const hedgeStrikeOffset = GetValueByName(sheetID, "IndexStranglesHedgeOffset", verbose);
  const hedgeDTE = GetValueByName(sheetID, "IndexStranglesHedgeDTE", verbose);
  var candidates= null;

  if (symbols && hedgeStrikeOffset && hedgeDTE)
  {
    candidates= GetIndexStrangleContractsSchwabByStrike(sheetID, symbols, hedgeDTE, hedgeStrikeOffset, verbose);
  }
  else
  {
    // Missing parameters
    candidates= [];
    Log(`Missing parameters: symbols= ${symbols}, DTE= ${hedgeDTE}, offset= ${hedgeStrikeOffset}`);
  }

  return candidates;
};