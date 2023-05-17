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
  const optionPrices= true;
  const verbose= false;
  var success= false;
  
  const nextUpdateTime= GetValueByName(sheetID, "IndexStranglesCandidatesUpdateNext", verbose);
  const currentTime= new Date();
  
  if (afterHours == undefined)
  {
    afterHours= true;
  }
  
  if (IsMarketOpen(sheetID, optionPrices, verbose) && (currentTime > nextUpdateTime)
        || currentTime.getDate() != nextUpdateTime.getDate() || afterHours)
  {
    // Only check during market hours or after a day change
    var candidates= [];
    var candidatesAdditional= [];
    const symbols= GetTableByNameSimple(sheetID, "IndexStranglesSymbols", verbose);
    const deltaCall= GetValueByName(sheetID, "IndexStranglesDeltaCall", verbose);
    const deltaPut= GetValueByName(sheetID, "IndexStranglesDeltaPut", verbose);
    const dte= GetValueByName(sheetID, "IndexStranglesDTE", verbose);
    
    candidates= GetIndexStrangleContracts(sheetID, symbols, dte, deltaCall, deltaPut, verbose);

    candidatesAdditional= GetIndexStrangleContracts(sheetID, symbols, 7, deltaCall, deltaPut, verbose);
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