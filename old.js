function RunTestOld()
{
  const one= 1;
  const two= 2;
  const three= 3;

  Log(`Testing... ${one}, $${two}, ${three}...`);
};


/**
 * RunBoxTradeCandidates()
 *
 * Find suitable contract expirations for box trades
 *
 */
function RunBoxTradeCandidatesOld(backupRun)
{
  // Declare constants and local variables
  const sheetID= GetMainSheetID();
  const verbose= false;
  var force= false;
  var success= false;
  
  const nextUpdateTime= GetValueByName(sheetID, "IndexStranglesBoxesUpdateNext", verbose);
  const currentTime= new Date();

  if (backupRun == undefined)
  {
    force= true;
  }
  
  if (force || (backupRun && currentTime > nextUpdateTime))
  {
    // Looks like we are due for an update
    const expirationTargets= GetTableByNameSimple(sheetID, "IndexStranglesBoxesDTEs", verbose);
    const underlying= GetValueByName(sheetID, "IndexStranglesBoxesUnderlying", verbose);
    var dteEarliest= 1000000;
    var dteLatest= 0;
    var expirations= null;

    if (expirationTargets && underlying)
    {
      // We have a list of target expirations
      for (var index in expirationTargets)
      {
        // Find expiration targets endpoints
        if (typeof expirationTargets[index][0] == "number")
        {
          if (dteLatest < expirationTargets[index][0])
          {
            dteLatest= expirationTargets[index][0]
          }
          if (dteEarliest > expirationTargets[index][0])
          {
            dteEarliest= expirationTargets[index][0]
          }
        }
      }

      if (dteEarliest <= dteLatest)
      {
        // Get the chain
        response= GetChainForSymbolByExpirationTDA(sheetID, underlying, dteEarliest, dteLatest, verbose);
      }
      else
      {
        Logger.log("[RunBoxTradeCandidates] Improper DTE endpoints: Earliest= %s, Latest= %s",
                    dteEarliest.toFixed(0), dteLatest.toFixed(0));
      }

      if (response)
      {
        // Data fetched -- extract
        expirations= ExtractExpirationsTDA(response, expirationTargets, verbose);
      }
      else
      {
        // Failed to fetch results
        Logger.log("[RunBoxTradeCandidates] Could not fetch option chain for <%s> for expirations between <%s> and <%s> days!",
                    underlying, dteEarliest.toFixed(0), dteLatest.toFixed(0));
      }

      if (expirations)
      {
        // We have viable expiration dates -- save them
        if (SetTableByName(sheetID, "IndexStranglesBoxesExpirations", expirations, verbose))
        {
          success= true;
          UpdateTime(sheetID, "IndexStranglesBoxesUpdateTime", verbose);
          SetValueByName(sheetID, "IndexStranglesBoxesUpdateStatus", "Updated [" + DateToLocaleString(currentTime) + "]", verbose);
        }
      }
    }
    else
    {
      Logger.log("[RunBoxTradeCandidates] Missing parameters: Underlying= %s, Expiration Targets= %s", underlying, expirationTargets);
    }
  }
  
  LogSend(sheetID);
  return success;
};