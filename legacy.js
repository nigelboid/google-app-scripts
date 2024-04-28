/**
 * GetContractsForSymbolByExpirationSchwab()
 *
 * Obtain option contracts for a given symbol and for a given expiration dates
 *
 */
function GetContractsForSymbolByExpirationSchwab(sheetID, underlying, dte, labelPuts, labelCalls, verbose)
{
  // Declare constants and local variables
  const url = ConstructUrlChainByExpirationsSchwab(underlying, dte, dte);;
  var contracts = null;
  
  if (url)
  {
    const chain = GetURLSchwab(sheetID, url, verbose);
    
    if (chain)
    {
      // Data fetched -- extract
      contracts= ExtractEarliestContractsSchwab(chain, labelPuts, labelCalls);
    }
    else
    {
      // Failed to fetch web pages
      LogThrottled(sheetID, "Could not fetch option chains!");
    }
  }
  else
  {
    // No chains to fetch?
    Log("Could not compile query!");
  }

  return contracts;
};


/**
 * ExtractEarliestContractsSchwab()
 *
 * Extract contract expiration dates from a list of returned contracts
 */
function ExtractEarliestContractsSchwab(chain, labelPuts, labelCalls)
{
  // Declare constants and local variables
  var labelDate = null;
  var expirations = null;
  var contractTypes = [labelPuts, labelCalls];
  var contractType = null;
  var contracts = {};
  
  while (contractType = contractTypes.shift())
  {
    // Extract data for each contract type (puts and calls)
    expirations = Object.keys(chain[contractType]).sort();
    
    if (expirations)
    {
      // We seem to have contract expiration dates
      while (expirations.length > 0)
      {
        // Find the earliest batch of contracts which satisfy our days-to-expiration constraint
        labelDate = expirations.shift();
      
        contracts[contractType] = {};
        contracts[contractType] = chain[contractType][labelDate];
        
        if (ValidateContracts(contracts[contractType]))
        {
          // Confirm we have valid contracts for this date to exit the search
          break;
        }
      }
    }
    else
    {
      Log("Data query returned no content");
      Log(contentParsed);
    }
  }
  
  return contracts;
};

/**
 * ValidateContracts()
 *
 * Confirm a contract in the list has valid data
 */
function ValidateContracts(contracts)
{
  // Declare constants and local variables
  var strikes= Object.keys(contracts);
  var badData= -999;
  var labelDelta= "delta";
  
  for (const strike of strikes)
  {
    // Search for at least one valid contract
    if (contracts[strike][0][labelDelta] != badData)
    {
      // Found one!
      return true;
    }
  }
  
  return false;
};


/**
 * ConstructUrlChainByExpirationsSchwab()
 *
 * Construct a URL to obtain a list of contracts with expiration dates
 */
function ConstructUrlChainByExpirationsSchwab(underlying, dteEarliest, dteSpan)
{
  // Declare constants and local variables
  const urlHead= "https://api.schwabapi.com/marketdata/v1/chains?";
  const urlSymbol= "symbol=" + underlying;
  const urlCount= "&strikeCount=500";
  const urlStrategy= "&strategy=SINGLE";
  var urlFromDate= "&fromDate=";
  var urlToDate= "&toDate=";
  
  // Construct dates in yyyy-mm-dd format by abusing the ISO string output method
  var dateParameter= new Date();
  
  // Aim for requested DTE
  dateParameter.setDate(dateParameter.getDate() + dteEarliest);
  urlFromDate+= dateParameter.toISOString().split('T').shift();
  
  // Collect expirations during the specified window (plus a buffer)
  dateParameter.setDate(dateParameter.getDate() + dteSpan);
  urlToDate+= dateParameter.toISOString().split('T').shift();
  
  return urlHead + urlSymbol + urlCount + urlStrategy + urlFromDate + urlToDate;
};


/**
 * GetIndexStrangleContractsSchwab()
 *
 * Obtain and select best matching strangles
 *
 */
function GetIndexStrangleContractsSchwabLegacy(sheetID, symbols, dte, deltaCall, deltaPut, verbose)
{
  // Declare constants and local variables
  const labelPuts= "putExpDateMap";
  const labelCalls= "callExpDateMap";
  var symbol= null;
  var contracts= null;
  var results= [];
  
  for (var index in symbols)
  {
    // Find candidates for each requested symbol
    symbol= symbols[index][0];
    
    if (symbol)
    {
      contracts= GetContractsForSymbolByExpirationSchwab(sheetID, symbol, dte, labelPuts, labelCalls, verbose);
    
      if (contracts)
      {
        // We have valid contracts -- sift through them for best delta matches
        if (contracts[labelPuts])
        {
          // Viable puts data
          results= results.concat(FindBestDeltaMatchSchwab(contracts[labelPuts], deltaPut, verbose));
        }
        
        if (contracts[labelCalls])
        {
          // Viable calls data
          results= results.concat(FindBestDeltaMatchSchwab(contracts[labelCalls], deltaCall, verbose));
        }
      }
      else
      {
        Log(`Failed to obtain a list of valid option contracts! (${contracts})`);
      }
    }
  }
  
  return results;
};


/**
 * FindBestDeltaMatchSchwab()
 *
 * Find a contract in the given set with closest delta value match
 *
 */
function FindBestDeltaMatchSchwab(contracts, deltaTarget, verbose)
{
  // Declare constants and local variables
  var strikes= Object.keys(contracts);
  var contractsBestMatch= {};
  var matches= [];
  var strike= null;
  var underlying= "";
  var deltaNew= 0;
  var deltaBest= 0;
  
  const labelSymbol= "symbol";
  const labelDelta= "delta";
  const labelLastPrice= "last";
  const labelDTE= "daysToExpiration";
  const labelContract= "contract";
  const symbolDelimiter= " ";
  const weekly= "W";
  const deltaTargetSensitivity= 0.01;
  
  deltaTarget/= 100;
  const deltaTargetMinimum= deltaTarget - deltaTargetSensitivity;
  const deltaTargetMaximum= deltaTarget + deltaTargetSensitivity;
  
  while (strike= strikes.shift())
  {
    // check each contract for a best match
    for (var contractDetails of contracts[strike])
    {
      // check each contract for this strike (e.g., AM and PM expiration for Index options
      if (contractDetails && contractDetails[labelSymbol] && contractDetails[labelDelta])
      {
        // we have viable details -- find the contract with the smallest delta near our target
        
        deltaNew= Math.abs(contractDetails[labelDelta]);
        if (deltaNew > deltaTargetMinimum && deltaNew < deltaTargetMaximum)
        {
          // within range
          
          underlying= contractDetails[labelSymbol].split(symbolDelimiter)[0];
          if (contractsBestMatch[underlying] == undefined)
          {
            // first hit within range -- create instance
            contractsBestMatch[underlying]= {};
            contractsBestMatch[underlying][labelContract]= {};
            contractsBestMatch[underlying][labelDelta]= 0;
          }
          
          deltaBest= Math.abs(contractsBestMatch[underlying][labelDelta]);
          
          // Choose the best match that is not too high (within one delta of the target) and otherwise closest to the target delta
          if ((deltaNew < deltaBest && deltaNew >= deltaTarget)
              || (deltaNew > deltaBest && deltaNew < deltaTargetMaximum && deltaBest < deltaTarget))
          {
            // found a better match!
            contractsBestMatch[underlying][labelContract]= contractDetails;
            contractsBestMatch[underlying][labelDelta]= contractDetails[labelDelta];
          }
        }
      }
    }
  }
  
  for (underlying in contractsBestMatch)
  {
    // Reformulate best matche as an array and add it to our list (prefer weeklies)
    if (underlying.endsWith(weekly) || contractsBestMatch[underlying.concat(weekly)] == undefined)
    {
      matches.push([contractsBestMatch[underlying][labelContract][labelSymbol],
                    contractsBestMatch[underlying][labelContract][labelLastPrice],
                    contractsBestMatch[underlying][labelContract][labelDelta],
                    contractsBestMatch[underlying][labelContract][labelDTE]
                  ]);
    }
  }
  
  return matches;
};/**
 * GetContractsForSymbolByExpirationSchwab()
 *
 * Obtain option contracts for a given symbol and for a given expiration dates
 *
 */
function GetContractsForSymbolByExpirationSchwab(sheetID, underlying, dte, labelPuts, labelCalls, verbose)
{
  // Declare constants and local variables
  const url = ConstructUrlChainByExpirationsSchwab(underlying, dte, dte);;
  var contracts = null;
  
  if (url)
  {
    const chain = GetURLSchwab(sheetID, url, verbose);
    
    if (chain)
    {
      // Data fetched -- extract
      contracts= ExtractEarliestContractsSchwab(chain, labelPuts, labelCalls);
    }
    else
    {
      // Failed to fetch web pages
      LogThrottled(sheetID, "Could not fetch option chains!");
    }
  }
  else
  {
    // No chains to fetch?
    Log("Could not compile query!");
  }

  return contracts;
};


/**
 * ExtractEarliestContractsSchwab()
 *
 * Extract contract expiration dates from a list of returned contracts
 */
function ExtractEarliestContractsSchwab(chain, labelPuts, labelCalls)
{
  // Declare constants and local variables
  var labelDate = null;
  var expirations = null;
  var contractTypes = [labelPuts, labelCalls];
  var contractType = null;
  var contracts = {};
  
  while (contractType = contractTypes.shift())
  {
    // Extract data for each contract type (puts and calls)
    expirations = Object.keys(chain[contractType]).sort();
    
    if (expirations)
    {
      // We seem to have contract expiration dates
      while (expirations.length > 0)
      {
        // Find the earliest batch of contracts which satisfy our days-to-expiration constraint
        labelDate = expirations.shift();
      
        contracts[contractType] = {};
        contracts[contractType] = chain[contractType][labelDate];
        
        if (ValidateContracts(contracts[contractType]))
        {
          // Confirm we have valid contracts for this date to exit the search
          break;
        }
      }
    }
    else
    {
      Log("Data query returned no content");
      Log(contentParsed);
    }
  }
  
  return contracts;
};

/**
 * ValidateContracts()
 *
 * Confirm a contract in the list has valid data
 */
function ValidateContracts(contracts)
{
  // Declare constants and local variables
  var strikes= Object.keys(contracts);
  var badData= -999;
  var labelDelta= "delta";
  
  for (const strike of strikes)
  {
    // Search for at least one valid contract
    if (contracts[strike][0][labelDelta] != badData)
    {
      // Found one!
      return true;
    }
  }
  
  return false;
};


/**
 * ConstructUrlChainByExpirationsSchwab()
 *
 * Construct a URL to obtain a list of contracts with expiration dates
 */
function ConstructUrlChainByExpirationsSchwab(underlying, dteEarliest, dteSpan)
{
  // Declare constants and local variables
  const urlHead= "https://api.schwabapi.com/marketdata/v1/chains?";
  const urlSymbol= "symbol=" + underlying;
  const urlCount= "&strikeCount=500";
  const urlStrategy= "&strategy=SINGLE";
  var urlFromDate= "&fromDate=";
  var urlToDate= "&toDate=";
  
  // Construct dates in yyyy-mm-dd format by abusing the ISO string output method
  var dateParameter= new Date();
  
  // Aim for requested DTE
  dateParameter.setDate(dateParameter.getDate() + dteEarliest);
  urlFromDate+= dateParameter.toISOString().split('T').shift();
  
  // Collect expirations during the specified window (plus a buffer)
  dateParameter.setDate(dateParameter.getDate() + dteSpan);
  urlToDate+= dateParameter.toISOString().split('T').shift();
  
  return urlHead + urlSymbol + urlCount + urlStrategy + urlFromDate + urlToDate;
};


/**
 * GetIndexStrangleContractsSchwab()
 *
 * Obtain and select best matching strangles
 *
 */
function GetIndexStrangleContractsSchwabLegacy(sheetID, symbols, dte, deltaCall, deltaPut, verbose)
{
  // Declare constants and local variables
  const labelPuts= "putExpDateMap";
  const labelCalls= "callExpDateMap";
  var symbol= null;
  var contracts= null;
  var results= [];
  
  for (var index in symbols)
  {
    // Find candidates for each requested symbol
    symbol= symbols[index][0];
    
    if (symbol)
    {
      contracts= GetContractsForSymbolByExpirationSchwab(sheetID, symbol, dte, labelPuts, labelCalls, verbose);
    
      if (contracts)
      {
        // We have valid contracts -- sift through them for best delta matches
        if (contracts[labelPuts])
        {
          // Viable puts data
          results= results.concat(FindBestDeltaMatchSchwab(contracts[labelPuts], deltaPut, verbose));
        }
        
        if (contracts[labelCalls])
        {
          // Viable calls data
          results= results.concat(FindBestDeltaMatchSchwab(contracts[labelCalls], deltaCall, verbose));
        }
      }
      else
      {
        Log(`Failed to obtain a list of valid option contracts! (${contracts})`);
      }
    }
  }
  
  return results;
};


/**
 * FindBestDeltaMatchSchwab()
 *
 * Find a contract in the given set with closest delta value match
 *
 */
function FindBestDeltaMatchSchwab(contracts, deltaTarget, verbose)
{
  // Declare constants and local variables
  var strikes= Object.keys(contracts);
  var contractsBestMatch= {};
  var matches= [];
  var strike= null;
  var underlying= "";
  var deltaNew= 0;
  var deltaBest= 0;
  
  const labelSymbol= "symbol";
  const labelDelta= "delta";
  const labelLastPrice= "last";
  const labelDTE= "daysToExpiration";
  const labelContract= "contract";
  const symbolDelimiter= " ";
  const weekly= "W";
  const deltaTargetSensitivity= 0.01;
  
  deltaTarget/= 100;
  const deltaTargetMinimum= deltaTarget - deltaTargetSensitivity;
  const deltaTargetMaximum= deltaTarget + deltaTargetSensitivity;
  
  while (strike= strikes.shift())
  {
    // check each contract for a best match
    for (var contractDetails of contracts[strike])
    {
      // check each contract for this strike (e.g., AM and PM expiration for Index options
      if (contractDetails && contractDetails[labelSymbol] && contractDetails[labelDelta])
      {
        // we have viable details -- find the contract with the smallest delta near our target
        
        deltaNew= Math.abs(contractDetails[labelDelta]);
        if (deltaNew > deltaTargetMinimum && deltaNew < deltaTargetMaximum)
        {
          // within range
          
          underlying= contractDetails[labelSymbol].split(symbolDelimiter)[0];
          if (contractsBestMatch[underlying] == undefined)
          {
            // first hit within range -- create instance
            contractsBestMatch[underlying]= {};
            contractsBestMatch[underlying][labelContract]= {};
            contractsBestMatch[underlying][labelDelta]= 0;
          }
          
          deltaBest= Math.abs(contractsBestMatch[underlying][labelDelta]);
          
          // Choose the best match that is not too high (within one delta of the target) and otherwise closest to the target delta
          if ((deltaNew < deltaBest && deltaNew >= deltaTarget)
              || (deltaNew > deltaBest && deltaNew < deltaTargetMaximum && deltaBest < deltaTarget))
          {
            // found a better match!
            contractsBestMatch[underlying][labelContract]= contractDetails;
            contractsBestMatch[underlying][labelDelta]= contractDetails[labelDelta];
          }
        }
      }
    }
  }
  
  for (underlying in contractsBestMatch)
  {
    // Reformulate best matche as an array and add it to our list (prefer weeklies)
    if (underlying.endsWith(weekly) || contractsBestMatch[underlying.concat(weekly)] == undefined)
    {
      matches.push([contractsBestMatch[underlying][labelContract][labelSymbol],
                    contractsBestMatch[underlying][labelContract][labelLastPrice],
                    contractsBestMatch[underlying][labelContract][labelDelta],
                    contractsBestMatch[underlying][labelContract][labelDTE]
                  ]);
    }
  }
  
  return matches;
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