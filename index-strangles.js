/**
 * RunIndexStranglesCandidates()
 *
 * Find suitable contracts based on a list of parameters
 *
 */
function RunIndexStranglesCandidates(afterHours)
{
  // Declare constants and local variables
  const sheetID= GetMainSheetID();
  const forceRefreshNowName= "IndexStranglesForceRefreshNow";
  const optionPrices= true;
  const verbose= false;
  var success= false;
  
  const forceRefreshNow= GetValueByName(sheetID, forceRefreshNowName, verbose);
  
  const nextUpdateTime= GetValueByName(sheetID, "IndexStranglesCandidatesUpdateNext", verbose);
  const currentTime= new Date();

  if (forceRefreshNow)
  {
    // User set a manual forced refresh flag for index strangles...
    if (verbose)
    {
      Logger.log("[RunIndexStranglesCandidates] Forcing a manual refresh of index strangles...");
      Logger.log("[RunIndexStranglesCandidates] Clearing the flag for manual refresh of index strangles...");
    }
    
    SetValueByName(sheetID, forceRefreshNowName, "", verbose);
  }
  
  if (afterHours == undefined)
  {
    afterHours= true;
  }
  
  if (forceRefreshNow || currentTime.getDate() != nextUpdateTime.getDate()
      || ((IsMarketOpen(sheetID, optionPrices, verbose) || afterHours) && (currentTime > nextUpdateTime)))
  {
    // Only check during market hours or after a day change
    var candidates= [];
    var candidatesAdditional= [];
    var dte= GetValueByName(sheetID, "IndexStranglesDTE", verbose);
    const symbols= GetTableByNameSimple(sheetID, "IndexStranglesSymbols", verbose);
    const deltaCall= GetValueByName(sheetID, "IndexStranglesDeltaCall", verbose);
    const deltaPut= GetValueByName(sheetID, "IndexStranglesDeltaPut", verbose);
    const firstCandidateIndex= 0;
    const dteIndex= 3;
    
    candidates= GetIndexStrangleContracts(sheetID, symbols, dte, deltaCall, deltaPut, verbose);

    if (candidates.length > 0)
    {
      // we have an initial set of candidates; adjust DTE for additional candidates
      dte= Math.max(candidates[firstCandidateIndex][dteIndex] + 1, 60);
    }
    candidatesAdditional= GetIndexStrangleContracts(sheetID, symbols, dte, deltaCall, deltaPut, verbose);
    
    if (candidatesAdditional.length > 0)
    {
      // Add a blank line
      candidates.push([""]);
    
      // Add near-term candidates to the list of our primary candidates
      candidates= candidates.concat(candidatesAdditional);
    }
    else
    {
      Logger.log("[RunIndexStranglesCandidates] Found no additional candidates <%s>!", candidatesAdditional);
    }

    if (candidates.length > 0)
    {
      // Looks like we obtained candidate contracts -- commit them to a table
      if (SetTableByName(sheetID, "IndexStranglesCandidates", candidates, verbose))
      {
        UpdateTime(sheetID, "IndexStranglesCandidatesUpdateTime", verbose);
        SetValueByName(sheetID, "IndexStranglesCandidatesUpdateStatus", "Updated [" + DateToLocaleString(currentTime) + "]", verbose);

        success= true;
      }
      else
      {
        SetValueByName(sheetID, "IndexStranglesCandidatesUpdateStatus",
                        "Failed to set new candidates [" + DateToLocaleString(currentTime) + "]", verbose);

        Logger.log(
          "[RunIndexStranglesCandidates] Failed to update table named 'IndexStranglesCandidates' in sheet <%s> with candidates: <%s>!",
          sheetID, candidates);
      }
    }
    else
    {
      // No candidates found?
      UpdateTime(sheetID, "IndexStranglesCandidatesUpdateTime", verbose);
      SetValueByName(sheetID, "IndexStranglesCandidatesUpdateStatus",
                      "Failed to find new candidates [" + DateToLocaleString(currentTime) + "]", verbose);

      Logger.log("[RunIndexStranglesCandidates] Found no candidates <%s>!", candidates);
      success= true;
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
  if (symbols && dte && deltaCall && deltaPut)
  {
    candidates= GetIndexStrangleContractsTDA(sheetID, symbols, dte, deltaCall, deltaPut, verbose);
  }
  else
  {
    // Missing parameters
    Logger.log("[GetIndexStrangleContracts] Missing parameters: symbols= %s, DTE= %s, calls delta= %s, puts delta= %s",
                symbols, dte, deltaCall, deltaPut);
  }

  return candidates;
};