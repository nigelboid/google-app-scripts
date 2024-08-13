/**
 * GetQuotesSchwab()
 *
 * Obtain quotes for a list of symbols
 *
 */
function GetQuotesSchwab(sheetID, symbols, labels, urlHead, verbose)
{
  // Declare constants and local variables
  const firstDataRow = 1;
  const symbolColumn = 0;
  var url = null;
  var urls = {};
  var symbolMap = null;
  var prices = null;
  
  if (symbols)
  {
    // We have symbols, proceed
    for (var vIndex = firstDataRow; vIndex < symbols.length; vIndex++)
    {
      // Compile a list of unique symbols
      if (symbols[vIndex][symbolColumn].length > 0)
      {
        // Live line
        url= ConstructUrlQuote(symbols[vIndex][symbolColumn], urlHead, verbose);
        if (url)
        {
          // Create entries for each URL in two mirroring maps for future reconciliation
          urls[symbols[vIndex][symbolColumn]] = url;
        }
      }
    }
    
    // Option quotes from Schwab require symbol remapping
    symbolMap = RemapSymbolsSchwab(sheetID, Object.keys(urls), verbose);
    url = ConstructUrlQuoteSchwab(Object.keys(symbolMap));

    if (url)
    {
      const quotes = GetURLSchwab(sheetID, url, verbose);
      
      if (quotes)
      {
        // Data fetched -- extract
        prices = ExtractPricesSchwab(quotes, symbolMap, urls, labels, verbose);
      }
      else
      {
        // Failed to fetch web pages
        LogThrottled(sheetID, "Could not fetch quotes!", verbose);
      }
    }
    else
    {
      // No prices to fetch?
      Log("Could not compile query!");
    }
  }
  else
  {
    // Missing parameters
    LogThrottled(sheetID, `Missing parameters: symbols= ${symbols}, headers= ${headers}`, verbose);
  }
  
  return prices;
};


/**
 * GetIndexStrangleContractsSchwabByDelta()
 *
 * Obtain and select best matching strangles by delta
 *
 */
function GetIndexStrangleContractsSchwabByDelta(sheetID, symbols, dte, deltaTargetCall, deltaTargetPut, verbose)
{
  // Declare constants and local variables
  const labelCall = "CALL";
  const labelPut = "PUT";
  const labelDateExpiration = "expirationDate";
  const labelSymbol = "symbol";
  const labelDelta = "delta";
  const labelPriceLast = "last";
  const labelDTE = "daysToExpiration";
  const preferredSettlement = "P";
  
  var contracts = [];
  
  for (const index in symbols)
  {
    // Find candidates for each requested underlying symbol
    const underlying = symbols[index][0];
    
    if (underlying)
    {
      const expirations = GetExpirationsByDTESchwab(sheetID, underlying, dte, verbose);
      if (expirations)
      {
        const dteList = Object.keys(expirations).sort(function(a, b) { return a - b; });
        var expirationDate = null;
        var contractCall = null;
        var contractPut = null;

        for (const index in dteList)
        {
          // Looks for candidates while we have expirations
          expirationDate = expirations[dteList[index]][labelDateExpiration];

          if (!contractCall)
          {
            // Still looking for an appropriate call
            contractCall = GetContractByDeltaSchwab
            (
              sheetID,
              underlying,
              expirationDate,
              labelCall,
              deltaTargetCall,
              preferredSettlement,
              verbose
            );
          }

          if (!contractPut)
          {
            // Still looking for an appropriate put
            contractPut = GetContractByDeltaSchwab
            (
              sheetID,
              underlying,
              expirationDate,
              labelPut,
              deltaTargetPut,
              preferredSettlement,
              verbose
            );
          }

          if (contractCall && contractPut)
          {
            // If we found both contracts, stop searching
            break;
          }
        }

        // Commit viable contracts
        if (contractCall)
        {
          contracts.push
          (
            [
              contractCall[labelSymbol],
              contractCall[labelPriceLast],
              contractCall[labelDelta],
              contractCall[labelDTE]
            ]
          );
        }
        if (contractPut)
        {
          contracts.push
          (
            [
              contractPut[labelSymbol],
              contractPut[labelPriceLast],
              contractPut[labelDelta],
              contractPut[labelDTE]
            ]
          );
        }
      }
    }
  }
  
  return contracts;
};


/**
 * GetIndexStrangleContractsSchwabByStrike()
 *
 * Obtain and select best matching strangles by strike
 *
 */
function GetIndexStrangleContractsSchwabByStrike(sheetID, symbols, dte, strikeOffset, verbose)
{
  // Declare constants and local variables
  const labelCall = "CALL";
  const labelPut = "PUT";
  const labelDateExpiration = "expirationDate";
  const labelSymbol = "symbol";
  const labelDelta = "delta";
  const labelPriceLast = "last";
  const labelDTE = "daysToExpiration";
  const preferredSettlement = "P";
  
  var contracts = [];
  
  for (const index in symbols)
  {
    // Find candidates for each requested underlying symbol
    const underlying = symbols[index][0];
    
    if (underlying)
    {
      const expirations = GetExpirationsByDTESchwab(sheetID, underlying, dte, verbose);
      if (expirations)
      {
        const dteList = Object.keys(expirations).sort(function(a, b) { return a - b; });
        var expirationDate = null;
        var contractCall = null;
        var contractPut = null;

        for (const index in dteList)
        {
          // Looks for candidates while we have expirations
          expirationDate = expirations[dteList[index]][labelDateExpiration];

          if (!contractCall)
          {
            // Still looking for an appropriate call
            contractCall = GetContractByStrikeSchwab
            (
              sheetID,
              underlying,
              expirationDate,
              labelCall,
              (1 + strikeOffset),
              preferredSettlement,
              verbose
            );
          }

          if (!contractPut)
          {
            // Still looking for an appropriate put
            contractPut = GetContractByStrikeSchwab
            (
              sheetID,
              underlying,
              expirationDate,
              labelPut,
              (1 - strikeOffset),
              preferredSettlement,
              verbose
            );
          }

          if (contractCall && contractPut)
          {
            // If we found both contracts, stop searching
            break;
          }
        }

        // Commit viable contracts
        if (contractCall)
        {
          contracts.push
          (
            [
              contractCall[labelSymbol],
              contractCall[labelPriceLast],
              contractCall[labelDelta],
              contractCall[labelDTE]
            ]
          );
        }
        if (contractPut)
        {
          contracts.push
          (
            [
              contractPut[labelSymbol],
              contractPut[labelPriceLast],
              contractPut[labelDelta],
              contractPut[labelDTE]
            ]
          );
        }
      }
    }
  }
  
  return contracts;
};


/**
 * GetExpirationsByDTESchwab()
 *
 * Obtain a list of expiration dates of contracts for a given underlying instrument
 *
 */
function GetExpirationsByDTESchwab(sheetID, underlying, dte, verbose)
{
  // Declare constants and local variables
  const url = ConstructUrlExpirationsSchwab(underlying);
  var expirationsByDTE = null;

  if (url)
  {
    const expirations = GetURLSchwab(sheetID, url, verbose);
    
    if (expirations)
    {
      // Data fetched -- extract
      expirationsByDTE = ExtractExpirationDatesByDTESchwab(expirations, dte, verbose);
    }
    else
    {
      // Failed to fetch web pages
      Log("Could not fetch expiration dates!");
    }
  }
  else
  {
    // No prices to fetch?
    Log("Could not compile query!");
  }

  return expirationsByDTE;
};


/**
 * GetContractByDeltaSchwab()
 *
 * Obtain the best matching contract by delta for a given underlying, expiration, and type
 *
 */
function GetContractByDeltaSchwab(sheetID, underlying, expirationDate, contractType, deltaTarget, preferredSettlement, verbose)
{
  // Declare constants and local variables
  const url = ConstructUrlChainByExpirationSchwab(underlying, expirationDate, contractType, verbose);
  var contract = null;

  const contractTypeMap =
  {
    "CALL" : "callExpDateMap",
    "PUT" : "putExpDateMap"
  }

  if (url)
  {
    const chain= GetURLSchwab(sheetID, url, verbose);
    
    if (chain)
    {
      // Data fetched -- extract
      for (const expirationLabel in chain[contractTypeMap[contractType]])
      {
        // Check each set of strikes (should only be one set!)
        contract = GetContractByBestDeltaMatchSchwab
        (
          chain[contractTypeMap[contractType]][expirationLabel],
          deltaTarget,
          preferredSettlement,
          verbose
        );

        if (contract)
        {
          // Just in case we somehow get more than one set of strikes, quit after finding a match
          break;
        }
      }
      
    }
    else
    {
      // Failed to fetch web pages
      Log(`Could not fetch option chain for ${expirationDate}`);
    }
  }
  else
  {
    // No prices to fetch?
    Log("Could not compile query!");
  }

  return contract;
};


/**
 * GetContractByStrikeSchwab()
 *
 * Obtain the best matching contract by strike for a given underlying, expiration, and type
 *
 */
function GetContractByStrikeSchwab(sheetID, underlying, expirationDate, contractType, strikeOffset, preferredSettlement, verbose)
{
  // Declare constants and local variables
  const url = ConstructUrlChainByExpirationSchwab(underlying, expirationDate, contractType, verbose);
  const labelUnderlyingPrice = "underlyingPrice";
  var targetStrike = null;
  var contract = null;

  const contractTypeMap =
  {
    "CALL" : "callExpDateMap",
    "PUT" : "putExpDateMap"
  }

  if (url)
  {
    const chain= GetURLSchwab(sheetID, url, verbose);
    
    if (chain)
    {
      // Data fetched -- compute target price
      targetStrike = chain[labelUnderlyingPrice] * strikeOffset;
      
      // Data fetched -- extract
      for (const expirationLabel in chain[contractTypeMap[contractType]])
      {
        // Check each set of strikes (should only be one set!)
        contract = GetContractByBestStrikeMatchSchwab
        (
          chain[contractTypeMap[contractType]][expirationLabel],
          targetStrike,
          preferredSettlement,
          verbose
        );

        if (contract)
        {
          // Just in case we somehow get more than one set of strikes, quit after finding a match
          break;
        }
      }
      
    }
    else
    {
      // Failed to fetch web pages
      Log(`Could not fetch option chain for ${expirationDate}`);
    }
  }
  else
  {
    // No prices to fetch?
    Log("Could not compile query!");
  }

  return contract;
};


/**
 * GetContractByBestDeltaMatchSchwab()
 *
 * Find the best matching contract by delta within the supplied chain
 *
 */
function GetContractByBestDeltaMatchSchwab(chain, deltaTarget, preferredSettlement, verbose)
{
  // Declare constants and local variables
  const deltaTargetSensitivity = 0.01;
  const labelSymbol = "symbol";
  const labelDelta = "delta";
  const labelSettlement = "settlementType";
  const valuePreferredSettlementType = "P";
  var delta = 0;
  var deltaPreferenceOffset = 0;
  var newSettlement = null;
  var currentSettlement = newSettlement;
  var contract = null;

  deltaTarget/= 100;
  const deltaTargetMinimum = deltaTarget - deltaTargetSensitivity;
  const deltaTargetMaximum = deltaTarget + deltaTargetSensitivity;

  if (preferredSettlement == undefined)
  {
    // No preference -- set to default (PM)
    preferredSettlement = valuePreferredSettlementType;
  }

  // Search through the chain for closest delta match within sensitivity bounds
  for (const strike in chain)
  {
    // Check each strike
    for (const index in chain[strike])
    {
      // Check each contract
      delta = Math.abs(chain[strike][index][labelDelta]);
      if (delta >= deltaTargetMinimum && delta <= deltaTargetMaximum)
      {
        if (contract)
        {
          // Determine a delta difference offset for unpreferred expiration types (e.g., AM expiration)
          newSettlement = chain[strike][index][labelSettlement];
          currentSettlement = contract[labelSettlement];

          if (newSettlement == preferredSettlement && currentSettlement != preferredSettlement)
          {
            // Negative offset tilts toward replacing unpreferred settlement type
            deltaPreferenceOffset = -valueSettlementdeltaPreferenceOffset;
            LogVerbose
            (
              `Prefer replacing contract <${contract[labelSymbol]}>: ` +
              `old delta = <${Math.abs(contract[labelDelta])}>, new delta = <${delta}>, ` +
              `old type = <${currentSettlement}>, new type = <${newSettlement}>, ` +
              `offset = <${deltaPreferenceOffset.toFixed(4)}>, new contract = <${chain[strike][index][labelSymbol]}>`,
              verbose
            );
          }
          else if (newSettlement != preferredSettlement && currentSettlement == preferredSettlement)
          {
            // Positive offset tilts toward keeping preferred settlement type
            deltaPreferenceOffset = valueSettlementdeltaPreferenceOffset;
            LogVerbose
            (
              `Prefer keeping <${contract[labelSymbol]}>: ` +
              `old delta = <${Math.abs(contract[labelDelta])}>, new delta = <${delta}>, ` +
              `old type = <${currentSettlement}>, new type = <${newSettlement}>, ` +
              `offset = <${deltaPreferenceOffset.toFixed(4)}>, new contract = <${chain[strike][index][labelSymbol]}>`,
              verbose
            );
          }
          else
          {
            deltaPreferenceOffset = 0;
            LogVerbose
            (
              `No preference for <${contract[labelSymbol]}>: ` +
              `old delta = <${Math.abs(contract[labelDelta])}>, new delta = <${delta}>, ` +
              `old type = <${currentSettlement}>, new type = <${newSettlement}>, ` +
              `offset = <${deltaPreferenceOffset.toFixed(4)}>, new contract = <${chain[strike][index][labelSymbol]}>`,
              verbose
            );
          }

          // Check for a potentially better match
          if ((Math.abs(delta - deltaTarget) + deltaPreferenceOffset) < Math.abs(Math.abs(contract[labelDelta]) - deltaTarget))
          {
            // Found a closer match!
            if (currentSettlement == preferredSettlement && newSettlement != preferredSettlement)
            {
              LogVerbose
              (
                `Found AM settlement type as a better delta match for contract <${contract[labelSymbol]}>: ` +
                `old delta = <${Math.abs(contract[labelDelta])}>, new delta = <${delta}>, ` +
                `old type = <${currentSettlement}>, new type = <${newSettlement}>, ` +
                `offset = <${deltaPreferenceOffset.toFixed(4)}>, new contract = <${chain[strike][index][labelSymbol]}>`,
                verbose
              );
            }
            else if (currentSettlement != preferredSettlement && newSettlement == preferredSettlement)
            {
              LogVerbose
              (
                `Found PM settlement type as a better delta match for contract <${contract[labelSymbol]}>: ` +
                `old delta = <${Math.abs(contract[labelDelta])}>, new delta = <${delta}>, ` +
                `old type = <${currentSettlement}>, new type = <${newSettlement}>, ` +
                `offset = <${deltaPreferenceOffset.toFixed(4)}>, new contract = <${chain[strike][index][labelSymbol]}>`,
                verbose
              );
            }
            else
            {
              LogVerbose
              (
                `Found a better delta match with the same settlement type for contract <${contract[labelSymbol]}>: ` +
                `old delta = <${Math.abs(contract[labelDelta])}>, new delta = <${delta}>, ` +
                `old type = <${currentSettlement}>, new type = <${newSettlement}>, ` +
                `offset = <${deltaPreferenceOffset.toFixed(4)}>, new contract = <${chain[strike][index][labelSymbol]}>`,
                verbose
              );
            }
            
            contract = chain[strike][index];
            LogVerbose(`Found a better match for <${chain[strike][index][labelSymbol]}>: delta = ${delta}`, verbose);
          }
          else
          {
            LogVerbose(`Ignoring a worse match: delta = ${delta}`, verbose);
          }
        }
        else
        {
          // First match!
          contract = chain[strike][index];
          LogVerbose(`Found our first match: delta = ${delta}`, verbose);
        }
      }
    }
  }

  return contract;
};


/**
 * GetContractByBestStrikeMatchSchwab()
 *
 * Find the best matching contract by delta within the supplied chain
 *
 */
function GetContractByBestStrikeMatchSchwab(chain, strikeTarget, preferredSettlement, verbose)
{
  // Declare constants and local variables
  const labelSettlement = "settlementType";
  const valuePreferredSettlementType = "P";
  var newSettlement = null;
  var contract = null;
  var strikeNearest = "";

  if (preferredSettlement == undefined)
  {
    // No preference -- set to default (PM)
    preferredSettlement = valuePreferredSettlementType;
  }


  // Search through the chain for closest delta match within sensitivity bounds
  for (const strike in chain)
  {
    // Check each strike
    if (Math.abs(strikeTarget - strikeNearest) > Math.abs(strikeTarget - strike))
    {
      // Found a closer target strike
      strikeNearest = strike;
    }
  }


  // Extract best contract for this strike price
  for (const index in chain[strikeNearest])
  {
    if (contract)
    {
      // Check for a better match
      if (chain[strikeNearest][index][labelSettlement] == preferredSettlement)
      {
        contract = chain[strikeNearest][index];
        LogVerbose(`Found our first match: strike = ${strikeNearest}`, verbose);
      }
    }
    else
    {
      // First match!
      contract = chain[strikeNearest][index];
      LogVerbose(`Found our first match: strike = ${strikeNearest}`, verbose);
    }
  }

  return contract;
};


/**
 * ExtractPricesSchwab()
 *
 * Extract pricing data from Schwab's JSON result
 *
 */
function ExtractPricesSchwab(quotes, symbolMap, urls, labels, verbose)
{
  // Define interesting quote parameters
  const labelQuote = "quote";
  const labelAssetType = "assetMainType";
  const labelSymbol = "symbol";
  const labelPriceChange = "netChange";
  const labelPriceLast = "lastPrice";
  const labelMark = "mark";
  const labelMarkChange = "markChange";
  const labelClosePrice = "closePrice";
  const labelBidPrice = "bidPrice";
  const labelAskPrice = "askPrice";
  const labelOptionDelta = "delta";
  const labelNAV = "nAV";
  const labelURL = "URL";
  const labelDebug = "Debug";
  const symbolFuturesCruftLength = 3;
  const valueNoData = "no data"
  
  // Define type exceptions
  const labelFuture = "FUTURE";
  const labelMutualFund = "MUTUAL_FUND";
  const labelOption = "OPTION";
  
  // Declare local variables
  var prices = {};
  var quoteSymbol = null;

  // Seed the default table column to quote parameter map
  const columnBid = "Bid";
  const columnAsk = "Ask";
  const columnLast = "Last";
  const columnClose = "Close";
  const columnChange = "Change";
  const columnDelta = "Delta";
  const columnMark = "Mark";
  var labelMap = {};

  // Default column label to quote label map
  labelMap[columnBid] = labelBidPrice;
  labelMap[columnAsk] = labelAskPrice;
  labelMap[columnLast] = labelPriceLast;
  labelMap[columnClose] = labelClosePrice;
  labelMap[columnChange] = labelPriceChange;
  labelMap[columnDelta] = labelOptionDelta;
  labelMap[columnMark] = labelMark;
  
  for (const quote in quotes)
  {
    // Process each returned quote
    if (quotes[quote][labelSymbol] != undefined)
    {
      // Symbol exists
      quoteSymbol = quotes[quote][labelSymbol];
      if (quotes[quote][labelAssetType] == labelFuture)
      {
        // Adjust futures symbols to remove expiration reference
        quoteSymbol = quoteSymbol.substring(0, quoteSymbol.length - symbolFuturesCruftLength);
        
        // Make a copy under the adjusted symbol for future reference
        quotes[quoteSymbol] = quotes[quote];
      }

      prices[symbolMap[quoteSymbol]] = {};
      prices[symbolMap[quoteSymbol]][labelURL] = urls[symbolMap[quoteSymbol]];
      
      if (labelDebug != undefined)
      {
        // debug activated -- commit raw data
        prices[symbolMap[quoteSymbol]][labelDebug] = JSON.stringify(quotes[quoteSymbol], null, 4);
      }
      
      // Adjust labelMap based quote specifics
      if (quotes[quote][labelAssetType] == labelMutualFund)
      {
        // Use NAV as last price for mutual funds
        labelMap[columnLast] = labelNAV;
      }
      else if (quotes[quote][labelAssetType] == labelOption)
      {
        // Use mark change as price change
        labelMap[columnChange] = labelMarkChange;
      }
      else
      {
        // Reset to defaults
        labelMap[columnLast] = labelPriceLast;
        labelMap[columnChange] = labelPriceChange;
      }
      
      for (const label in labels)
      {
        // Pull a value for each desired label
        if (quotes[quote][labelQuote][labelMap[labels[label]]] == undefined)
        {
          // No data for this quote item
          prices[symbolMap[quoteSymbol]][labels[label]] = valueNoData;
        }
        else
        {
          // Data obtained -- translate and commit to our storage
          prices[symbolMap[quoteSymbol]][labels[label]] = quotes[quoteSymbol][labelQuote][labelMap[labels[label]]];
        }
      }
    }
  }
  
  return prices;
};


/**
 * ExtractExpirationDatesByDTESchwab()
 *
 * Extract a list of dates from Schwab's JSON reply limited to a DTE range
 *
 */
function ExtractExpirationDatesByDTESchwab(expirations, minimumDaysToExpiration, verbose)
{
  const expirationsListLabel = "expirationList";
  const daysToExpirationLabel = "daysToExpiration";
  var daysToExpiration = null;
  var expirationDates = {};

  for (var index in expirations[expirationsListLabel])
  {
    daysToExpiration = expirations[expirationsListLabel][index][daysToExpirationLabel];
    if (daysToExpiration >= minimumDaysToExpiration)
    {
      // Found a viable expiration -- copy it to our list
      expirationDates[daysToExpiration] = expirations[expirationsListLabel][index];
    }
  }

  return expirationDates;
};


/**
 * RemapSymbolsSchwab()
 *
 * Remap symbols to the Schwab format
 *   covers: options, indexes, futures, special cases (e.g., US10Y)
 *
 */
function RemapSymbolsSchwab(sheetID, symbols, verbose)
{
  // Declare constants and local variables
  const optionDetailsLength = "YYMMDDT00000000".length;
  const symbolMapProvided = GetParameters(sheetID, "ParameterMapSymbols", verbose);
  const symbolOptionUnderlyingPadding = 6;
  const symbolOptionDateStep = 2;
  const symbolOptionTypeStep = 1;

  var symbolMap = {};
  var symbolSchwab = "";
  var underlying = "";
  var date = "";
  var month = "";
  var year = "";
  var type = "";
  var strike = "";
  var symbolIndex = 0;

  for (const quoteSymbol of symbols)
  {
    if (symbolMapProvided[quoteSymbol] != undefined)
    {
      // Create our symbol map
      symbolMap[symbolMapProvided[quoteSymbol.toUpperCase()]] = quoteSymbol;
    }
    else if (quoteSymbol.length > optionDetailsLength)
    {
      // Re-map each apparent option symbol
    
      underlying = quoteSymbol.slice(0, -optionDetailsLength);
      
      symbolIndex = underlying.length;
      year = quoteSymbol.slice(symbolIndex, symbolIndex+= symbolOptionDateStep);
      month = quoteSymbol.slice(symbolIndex, symbolIndex+= symbolOptionDateStep);
      date = quoteSymbol.slice(symbolIndex, symbolIndex+= symbolOptionDateStep);
      
      type = quoteSymbol.slice(symbolIndex, symbolIndex+= symbolOptionTypeStep);
      
      strike = quoteSymbol.slice(symbolIndex);
      
      symbolSchwab = underlying.padEnd(symbolOptionUnderlyingPadding, " ") + year + month + date + type + strike;
      symbolMap[symbolSchwab] = quoteSymbol;
    }
    else
    {
      // Preserve all other symbols
      symbolMap[quoteSymbol] = quoteSymbol;
    }
  }
  
  return symbolMap;
};


/**
 * ConstructUrlQuoteSchwab()
 *
 * Construct a Schwab quote query URL for specified symbols
 *
 */
function ConstructUrlQuoteSchwab(symbols)
{
  // Declare constants and local variables
  const urlHead = "https://api.schwabapi.com/marketdata/v1/quotes?";
  const urlFields = "fields=quote";
  const urlIndicative = "&indicative=false";
  const urlSymbols = "&symbols=" + symbols.join(",");
  
  return urlHead + urlFields + urlIndicative + urlSymbols;
};


/**
 * ConstructUrlChainByExpirationSchwab()
 *
 * Construct a URL to obtain a list of contracts for a specific expiration date
 */
function ConstructUrlChainByExpirationSchwab(underlying, expirationDate, contractType, verbose)
{
  // Declare constants and local variables
  const urlHead = "https://api.schwabapi.com/marketdata/v1/chains?";
  const urlSymbol = "symbol=" + underlying;
  const urlCount = "&strikeCount=500";
  const urlStrategy = "&strategy=SINGLE";
  const urlFromDate = "&fromDate=" + expirationDate;
  const urlToDate = "&toDate=" + expirationDate;
  var urlContractType = "&contractType=";
  
  // Validate cotnact type parameter
  if (["ALL", "PUT", "CALL"].includes(contractType))
  {
    // Valid cotnract type -- set
    urlContractType += contractType;
  }
  else
  {
    // Unrecognized contract type -- set to default (puts and calls)
    urlContractType += "ALL";
    LogVerbose("Set contract type to default (all: puts and calls).", verbose)
  }
  
  return urlHead + urlSymbol + urlCount + urlStrategy + urlFromDate + urlToDate + urlContractType;
};


/**
 * ConstructUrlExpirationsSchwab()
 *
 * Construct a URL to obtain a list of expiration dates for a given underlying
 */
function ConstructUrlExpirationsSchwab(underlying)
{
  // Declare constants and local variables
  const urlHead = "https://api.schwabapi.com/marketdata/v1/expirationchain?";
  const urlSymbol = "symbol=" + underlying;
  
  return urlHead + urlSymbol;
};


/**
 * ComposeHeadersGetSchwab()
 *
 * Compose required headers for our requests
 */
function ComposeHeadersGetSchwab(sheetID, verbose)
{
  // Declare constants and local variables
  const accessToken = GetAccessTokenSchwab(sheetID, verbose);
  var headers = null;
  
  if (accessToken)
  {
    // We have an access token, proceed
    headers =
    {
      "Accept" : "application/json",
      "Schwab-Client-CorrelId": "ISLE",
      "Authorization" : "Bearer " + accessToken
    };
  }
  else
  {
    // We did not get an access token!
    LogThrottled(sheetID, `Cannot request quotes without a valid access token [${accessToken}]!`, verbose);
  }
  
  return headers;
};


/**
 * GetAccessTokenSchwab()
 *
 * Obtain a fresh and valid access token for Schwab API queries
 *
 */
function GetAccessTokenSchwab(sheetID, verbose)
{
  // Declare constants and local variables
  const currentTime = new Date();
  var accessTokenExpirationTime = GetValueByName(sheetID, "ParameterSchwabTokenAccessExpiration", verbose);
  var accessToken = GetValueByName(sheetID, "ParameterSchwabTokenAccess", verbose);
  
  if (!accessToken || !accessTokenExpirationTime || (currentTime > accessTokenExpirationTime))
  {
    // Invalidate the expired or unstamped access token
    LogVerbose(`Refreshing invalid access token (expiration stamp ${accessTokenExpirationTime})...`, verbose);
    LogVerbose(`Old access token: ${accessToken}`, verbose);
    accessToken = null;
  
    // Attempt to refresh an invalid access token
    // const refreshToken = GetRefreshTokenSchwab(sheetID, verbose);
    const refreshToken = GetValueByName(sheetID, "ParameterSchwabTokenRefresh", verbose);

    // Request a new access token using our valid refreh token
    if (refreshToken)
    {
      // Use the refresh token to obrain a new access token
      const url = "https://api.schwabapi.com/v1/oauth/token";
      const payload = {
        'grant_type': 'refresh_token',
        'refresh_token': refreshToken
      };

      const content = PostURLSchwab(sheetID, url, payload, verbose);

      if (content)
      {
        const accessTokenTTL = content["expires_in"];
        const accessTokenTTLOffset = 60 * 5;

        accessToken = content["access_token"];
        
        if (accessToken && accessTokenTTL)
        {
          // Preserve the new Access Token and its expiration time
          accessTokenExpirationTime = new Date();
          accessTokenExpirationTime.setSeconds(accessTokenExpirationTime.getSeconds() + accessTokenTTL - accessTokenTTLOffset);
          
          SetValueByName(sheetID, "ParameterSchwabTokenAccess", accessToken, verbose);
          SetValueByName(sheetID, "ParameterSchwabTokenAccessExpiration", accessTokenExpirationTime, verbose);
        }
        else
        {
          LogThrottled(sheetID, `Failed to obtain refreshed access token [${accessToken}] and its time-to-live [${accessTokenTTL}]!`);
        }
      }
    }
    else
    {
      const refreshTokenStaleLoggingThrottle = 60 * 60 * 24;

      LogThrottled(sheetID, `Invalid refresh token <${refreshToken}>!`, verbose, refreshTokenStaleLoggingThrottle);
    }
  }
  
  return accessToken;
};


/**
 * GetURLSchwab()
 *
 * Fetch the supplied Schwab API URL via GET
 */
function GetURLSchwab(sheetID, url, verbose)
{
  const headers = ComposeHeadersGetSchwab(sheetID, verbose);
  const payload = null;
  var response = null;
  var content = null;
  
  if (headers)
  {
    response = FetchURLSchwab(sheetID, url, headers, "get", payload, verbose);
  }
  else
  {
    LogThrottled(sheetID, `Missing parameters: headers= ${headers}`, verbose);
  }

  if (response)
  {
    content = ExtractContentSchwab(sheetID, response);
  }
  else
  {
    LogThrottled(sheetID, `Received no response for query  <${url}>`, verbose);
  }

  return content;
};


/**
 * PostURLSchwab()
 *
 * Fetch the supplied Schwab API URL via POST
 */
function PostURLSchwab(sheetID, url, payload, verbose)
{
  const key= GetValueByName(sheetID, "ParameterSchwabKey", verbose);
  var headers = null;
  var response = null;
  var content = null;

  if (key)
  {
    headers =
    {
      'Authorization': "Basic " + Utilities.base64Encode(key)
    };

    response = FetchURLSchwab(sheetID, url, headers, "post", payload, verbose);
  }
  else
  {
    LogThrottled(sheetID, `Could not read stored key: ${key}`, verbose);
  }

  if (response)
  {
    content = ExtractContentSchwab(sheetID, response);
  }
  else
  {
    LogThrottled(sheetID, `Received no response for query  <${url}>`, verbose);
  }

  return content;
};


/**
 * FetchURLSchwab()
 *
 * Fetch the supplied Schwab API URL
 */
function FetchURLSchwab(sheetID, url, headers, method, payload, verbose)
{
  var response = null;
  var options = {};

  if (headers)
  {
    options["headers"] = headers;
  }

  if (method)
  {
    options["method"] = method;
  }
  else
  {
    options["method"] = "get";
  }

  if (payload)
  {
    options["payload"] = payload;
  }

  options["muteHttpExceptions"] = true;
  
  try
  {
    response = UrlFetchApp.fetch(url, options);
  }
  catch (error)
  {
    LogThrottled(sheetID, error.message, verbose);
  }

  return response;
};


/**
 * ExtractContentSchwab()
 *
 * Extract parsed JSON payload from a Schwab response
 */
function ExtractContentSchwab(sheetID, response)
{
  // Declare constants and local variables
  const responseOK = 200;
  const contentParsed = JSON.parse(response.getContentText());
  
  if (response.getResponseCode() != responseOK)
  {
    var errorMessage = `Data query returned error code <${response.getResponseCode()}>`;
    
    if (contentParsed)
    {
      errorMessage = errorMessage + ":\n\n" + JSON.stringify(contentParsed, null, 4);
    }
    else
    {
      errorMessage += "!";
    }

    LogThrottled(sheetID, errorMessage);
  }

  return contentParsed;
};