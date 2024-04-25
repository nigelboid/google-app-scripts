/**
 * GetQuotesSchwab()
 *
 * Obtain quotes for a list of symbols
 *
 */
function GetQuotesSchwab(sheetID, symbols, labels, urlHead, verbose)
{
  // Declare constants and local variables
  var firstDataRow= 1;
  var symbolColumn= 0;
  var url= null;
  var urls= {};
  var symbolMap= null;
  var prices= null;
  
  if (symbols)
  {
    // We have symbols, proceed
    for (var vIndex= firstDataRow; vIndex < symbols.length; vIndex++)
    {
      // Compile a list of unique symbols
      if (symbols[vIndex][symbolColumn].length > 0)
      {
        // Live line
        url= ConstructUrlQuote(symbols[vIndex][symbolColumn], urlHead, verbose);
        if (url)
        {
          // Create entries for each URL in two mirroring maps for future reconciliation
          urls[symbols[vIndex][symbolColumn]]= url;
        }
      }
    }
    
    // Option quotes from Schwab require symbol remapping
    symbolMap= RemapSymbolsSchwab(sheetID, Object.keys(urls), verbose);
    url= ConstructUrlQuoteSchwab(Object.keys(symbolMap));

    if (url)
    {
      const quotes= GetURLSchwab(sheetID, url, verbose);
      
      if (quotes)
      {
        // Data fetched -- extract
        prices= ExtractPricesSchwab(quotes, symbolMap, urls, labels, verbose);
      }
      else
      {
        // Failed to fetch web pages
        Log("Could not fetch quotes!");
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
    LogThrottled(sheetID, `Missing parameters: symbols= ${symbols}, headers= ${headers}`);
  }
  
  return prices;
};


/**
 * ExtractPricesSchwab()
 *
 * Extract pricing data from Schwab's JSON result
 *
 */
function ExtractPricesSchwab(quotes, symbolMap, urls, labels, verbose)
{
  // Declare constants and local variables
  var prices= {};
  var quoteSymbol= null;
  
  if (quotes)
  {
    // We have quotes -- process them
    
    // First, define interesting quote parameters
    const labelQuote= "quote";
    const labelAssetType= "assetMainType";
    const labelSymbol= "symbol";
    const labelChangePrice= "netChange";
    const labelLastPrice= "lastPrice";
    const labelClosePrice= "closePrice";
    const labelBidPrice= "bidPrice";
    const labelAskPrice= "askPrice";
    const labelOptionDelta= "delta";
    const labelNAV= "nAV";
    const labelURL= "URL";
    const labelDebug= "Debug";
    const symbolFuturesCruftLength= 3;

    // Second, seed the default table column to quote parameter map
    const columnBid= "Bid";
    const columnAsk= "Ask";
    const columnLast= "Last";
    const columnClose= "Close";
    const columnChange= "Change";
    const columnDelta= "Delta";
    var labelMap= {};
    labelMap[columnBid]= labelBidPrice;
    labelMap[columnAsk]= labelAskPrice;
    labelMap[columnLast]= labelLastPrice;
    labelMap[columnClose]= labelClosePrice;
    labelMap[columnChange]= labelChangePrice;
    labelMap[columnDelta]= labelOptionDelta;
    
    // Lastly, define type exceptions
    const typeMutualFund= "MUTUAL_FUND";
    const typeFuture= "FUTURE";
    
    for (const quote in quotes)
    {
      // Process each returned quote
      if (quotes[quote][labelSymbol] != undefined)
      {
        // Symbol exists
        quoteSymbol= quotes[quote][labelSymbol];
        if (quotes[quote][labelAssetType] == typeFuture)
        {
          // Adjust futures symbols to remove expiration reference
          quoteSymbol= quoteSymbol.substring(0, quoteSymbol.length - symbolFuturesCruftLength);
          
          // Make a copy under the adjusted symbol for future reference
          quotes[quoteSymbol]= quotes[quote];
        }

        prices[symbolMap[quoteSymbol]]= {};
        prices[symbolMap[quoteSymbol]][labelURL]= urls[symbolMap[quoteSymbol]];
        
        if (labelDebug != undefined)
        {
          // debug activated -- commit raw data
          prices[symbolMap[quoteSymbol]][labelDebug]= JSON.stringify(quotes[quoteSymbol], null, 4);
        }
        
        // Adjust labelMap based quote specifics
        if (quotes[quote][labelAssetType] == typeMutualFund)
        {
          // Use NAV as last price for mutual funds
          labelMap[columnLast]= labelNAV;
          labelMap[columnChange]= labelChangePrice;
          labelMap[columnClose]= labelClosePrice;
        }
        else
        {
          // Set to defaults
          labelMap[columnLast]= labelLastPrice;
          labelMap[columnChange]= labelChangePrice;
          labelMap[columnClose]= labelClosePrice;
        }
        
        for (const label in labels)
        {
          // Pull a value for each desired label
          if (quotes[quote][labelQuote][labelMap[labels[label]]] == undefined)
          {
            // No data for this quote item
            prices[symbolMap[quoteSymbol]][labels[label]]= "no data";
          }
          else
          {
            // Data obtained -- translate and commit to our storage
            prices[symbolMap[quoteSymbol]][labels[label]]= quotes[quoteSymbol][labelQuote][labelMap[labels[label]]];
          }
        }
      }
    }
  }
  else
  {
    Log("Could not obtain quotes!");
  }
  
  return prices;
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
  const optionDetailsLength= "YYMMDDT00000000".length;
  const symbolMapProvided= GetParameters(sheetID, "ParameterMapSymbols", verbose);
  const symbolOptionUnderlyingPadding= 6;
  const symbolOptionDateStep= 2;
  const symbolOptionTypeStep= 1;

  var symbolMap= {};
  var symbolSchwab= "";
  var underlying= "";
  var date= "";
  var month= "";
  var year= "";
  var type= "";
  var strike= "";
  var symbolIndex= 0;

  for (const quoteSymbol of symbols)
  {
    if (symbolMapProvided[quoteSymbol] != undefined)
    {
      // Create our symbol map
      symbolMap[symbolMapProvided[quoteSymbol.toUpperCase()]]= quoteSymbol;
    }
    else if (quoteSymbol.length > optionDetailsLength)
    {
      // Re-map each apparent option symbol
    
      underlying= quoteSymbol.slice(0, -optionDetailsLength);
      
      symbolIndex= underlying.length;
      year= quoteSymbol.slice(symbolIndex, symbolIndex+= symbolOptionDateStep);
      month= quoteSymbol.slice(symbolIndex, symbolIndex+= symbolOptionDateStep);
      date= quoteSymbol.slice(symbolIndex, symbolIndex+= symbolOptionDateStep);
      
      type= quoteSymbol.slice(symbolIndex, symbolIndex+= symbolOptionTypeStep);
      
      strike= quoteSymbol.slice(symbolIndex);
      
      symbolSchwab= underlying.padEnd(symbolOptionUnderlyingPadding, " ") + year + month + date + type + strike;
      symbolMap[symbolSchwab]= quoteSymbol;
    }
    else
    {
      // Preserve all other symbols
      symbolMap[quoteSymbol]= quoteSymbol;
    }
  }
  
  return symbolMap;
};


/**
 * GetIndexStrangleContractsSchwab()
 *
 * Obtain and select best matching strangles
 *
 */
function GetIndexStrangleContractsSchwab(sheetID, symbols, dte, deltaCall, deltaPut, verbose)
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
 * GetContractsForSymbolByExpirationSchwab()
 *
 * Obtain option contract expiration dates for a given symbol
 *
 */
function GetContractsForSymbolByExpirationSchwab(sheetID, symbol, dte, labelPuts, labelCalls, verbose)
{
  // Declare constants and local variables
  var contracts= {};
  const response= GetChainForSymbolByExpirationSchwab(sheetID, symbol, dte, dte, verbose);
  
  if (response)
  {
    // Data fetched -- extract
    contracts= ExtractEarliestContractsSchwab(response, labelPuts, labelCalls);
  }
  else
  {
    // Failed to fetch results
    LogThrottled(sheetID, "Could not fetch option chains!");
    LogThrottled(sheetID, `Response: ${response}`);
  }

  return contracts;
};


/**
 * GetChainForSymbolByExpirationSchwab()
 *
 * Obtain option contract expiration dates for a given symbol
 *
 */
function GetChainForSymbolByExpirationSchwab(sheetID, symbol, dteEarliest, dteSpan, verbose)
{
  // Declare constants and local variables
  var headers= ComposeHeadersSchwab(sheetID, verbose);
  var url= null;
  var response= null;
  
  if (headers)
  {
    // We have viable headers and query parameters, proceed
    url= ConstructUrlChainByExpirationsSchwab(symbol, dteEarliest, dteSpan);

    if (verbose)
    {
      Log(`Query: ${url}`);
    }

    if (url)
    {
      // We have a URL -- get it!
      try
      {
        response= UrlFetchApp.fetch(url, {'headers' : headers});
      }
      catch (error)
      {
        LogThrottled(sheetID, error.message, verbose);
      }
    }
    else
    {
      // No prices to fetch?
      Log("Could not compile query.");
    }
  }
  else
  {
    // Missing parameters
    LogThrottled(sheetID, `Missing parameters: headers= ${headers}`);
  }

  return response;
}


/**
 * ExtractEarliestContractsSchwab()
 *
 * Extract contract expiration dates from a list of returned contracts
 */
function ExtractEarliestContractsSchwab(response, labelPuts, labelCalls)
{
  // Declare constants and local variables
  var contentParsed= ExtractContentSchwab(response);
  var labelDate= null;
  var expirations= [];
  var contractTypes= [labelPuts, labelCalls];
  var contractType= null;
  var contracts= {};
  
  if (contentParsed)
  {
    while (contractType= contractTypes.shift())
    {
      // Extract data for each contract type (puts and calls)
      expirations= Object.keys(contentParsed[contractType]).sort();
      
      if (expirations)
      {
        // We seem to have contract expiration dates
        while (expirations.length > 0)
        {
          // Find the earliest batch of contracts which satisfy our days-to-expiration constraint
          labelDate= expirations.shift();
        
          contracts[contractType]= {};
          contracts[contractType]= contentParsed[contractType][labelDate];
          
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
  }
  else
  {
    Log("Query returned no data!");
  }
  
  return contracts;
};


/**
 * ExtractExpirationsSchwab()
 *
 * Extract matching contract expiration dates from an option chain
 */
function ExtractExpirationsSchwab(response, expirationTargets, verbose)
{
  // Declare constants and local variables
  const labelPuts= "putExpDateMap";
  const contentParsed= ExtractContentSchwab(response);
  
  if (contentParsed)
  {
    // Looks like we have a valid data response
    const dateDelimiter= ":";
    var expirations= Object.keys(contentParsed[labelPuts]).sort();
    var expirationsMapped= [];
    var dte= null;

    if (expirations)
    {
      // We have a list of expirations -- now map them to desired DTE targets
      for (var target in expirationTargets)
      {
        // Match an expiration date to each valid target
        if (typeof expirationTargets[target][0] == "number")
        {
          for (var expiration of expirations)
          {
            // Find the earliest expiration date which satisfies our days-to-expiration constraint
            dte= expiration.split(dateDelimiter)[1];

            if( dte >= expirationTargets[target][0])
            {
              // We found the earliest expiration date
              const labelSymbol= "symbol";
              const symbolDelimiter= " ";
              const weekly= "W";
              const strike= Object.keys(contentParsed[labelPuts][expiration])[0];
              var underlying= "";

              // Determine the udnerlying symbol from a given quote
              for (var quote in contentParsed[labelPuts][expiration][strike])
              {
                underlying= contentParsed[labelPuts][expiration][strike][quote][labelSymbol].split(symbolDelimiter)[0];
                if (underlying.endsWith(weekly))
                {
                  // Prefer weeklies
                  break;
                }
              }

              expirationsMapped[target]= [expiration.split(dateDelimiter)[0], underlying];
              break;
            }
          }
        }
      }
    }
    else
    {
      Log(`Data query returned no viable expirations: <${expirations}>`);
      Log(contentParsed);
    }
  }
  else
  {
    Log("Query returned no data!");
  }

  return expirationsMapped;
};


/**
 * ValidateContracts()
 *
 * Confrim a contract in the list has valid data
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
 * ConstructUrlQuoteSchwab()
 *
 * Construct a Schwab quote query URL for specified symbols
 *
 */
function ConstructUrlQuoteSchwab(symbols)
{
  // Declare constants and local variables
  const urlHead= "https://api.schwabapi.com/marketdata/v1/quotes?";
  const urlFields= "fields=quote";
  const urlIndicative= "&indicative=false";
  const urlSymbols= "&symbols=" + symbols.join(",");
  
  return urlHead + urlFields + urlIndicative + urlSymbols;
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
 * ConstructUrlChainByExpirationSchwab()
 *
 * Construct a URL to obtain a list of contracts for a specific expiration date
 */
function ConstructUrlChainByExpirationSchwab(underlying, expiration, contractType, verbose)
{
  // Declare constants and local variables
  const urlHead= "https://api.schwabapi.com/marketdata/v1/chains?";
  const urlSymbol= "symbol=" + underlying;
  const urlCount= "&strikeCount=500";
  const urlStrategy= "&strategy=SINGLE";
  var urlFromDate= "&fromDate=" + expiration;
  var urlToDate= "&toDate=" + expiration;
  var urlContractType= "&contractType=";
  
  // Validate cotnact type parameter
  if (!["ALL", "PUT", "CALL"].includes(contractType))
  {
    // Set to default (puts and calls)
    urlContractType+= "ALL";
    LogVerbose("Set contract type to default (ALL).", verbose)
  }
  else
  {
    urlContractType+= contractType;
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
  const urlHead= "https://api.schwabapi.com/marketdata/v1/expirationchain?";
  const urlSymbol= "symbol=" + underlying;
  
  return urlHead + urlSymbol;
};


/**
 * ComposeHeadersSchwab()
 *
 * Compose required headers for our requests
 */
function ComposeHeadersSchwab(sheetID, verbose)
{
  // Declare constants and local variables
  const accessToken= GetAccessTokenSchwab(sheetID, verbose);
  var headers= null;
  
  if (accessToken)
  {
    // We have an access token, proceed
    headers= {
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
  const currentTime= new Date();
  var accessTokenExpirationTime= GetValueByName(sheetID, "ParameterSchwabTokenAccessExpiration", verbose);
  var accessToken= GetValueByName(sheetID, "ParameterSchwabTokenAccess", verbose);
  
  if (!accessToken || !accessTokenExpirationTime || (currentTime > accessTokenExpirationTime))
  {
    // Invalidate the expired or unstamped access token
    LogVerbose(`Refreshing invalid access token (expiration stamp ${accessTokenExpirationTime})...`, verbose);
    LogVerbose(`Old access token: ${accessToken}`, verbose);
    accessToken= null;
  
    // Attempt to refresh an invalid access token
    const key= GetValueByName(sheetID, "ParameterSchwabKey", verbose);
    const refreshToken= GetRefreshTokenSchwab(sheetID, verbose);
    var response= null;

    // Request a new access token using our valid refreh token
    if (key && refreshToken)
    {
      // we have necessary parameters, proceed to obtain token
      const formData= {
        'grant_type': 'refresh_token',
        'refresh_token': refreshToken
      };

      const headers = {
        'Authorization': "Basic " + Utilities.base64Encode(key)
      };
      
      const options = {
        'method' : 'post',
        'headers' : headers,
        'payload' : formData
      };
      
      try
      {
        response= UrlFetchApp.fetch("https://api.schwabapi.com/v1/oauth/token", options);
        
        if (response)
        {
          // Looks like we have a valid data response
          const keyAccessToken= "access_token";
          const keyAccessTokenTTL= "expires_in";
          const contentParsed= ExtractContentSchwab(response);
          
          if (contentParsed)
          {
            const accessTokenTTL= contentParsed[keyAccessTokenTTL];
            const accessTokenTTLOffset= 60 * 5;

            accessToken= contentParsed[keyAccessToken];
            
            if (accessToken && accessTokenTTL)
            {
              // Preserve the new Access Token and its expiration time
              accessTokenExpirationTime= new Date();
              accessTokenExpirationTime.setSeconds(accessTokenExpirationTime.getSeconds() + accessTokenTTL - accessTokenTTLOffset);
              
              SetValueByName(sheetID, "ParameterSchwabTokenAccess", accessToken, verbose);
              SetValueByName(sheetID, "ParameterSchwabTokenAccessExpiration", accessTokenExpirationTime, verbose);
            }
            else
            {
              Log(`Failed to obtain refreshed access token [${accessToken}] and its time-to-live [${accessTokenTTL}]!`);
            }
          }
          else
          {
            LogThrottled(sheetID, "Received error while trying to refresh access token.");
          }
        }
        else
        {
          LogThrottled(sheetID, "Received no response while trying to refresh access token.");
        }
      }
      catch (error)
      {
        LogThrottled(sheetID, error.message, verbose);
      }
    }
    else if (verbose)
    {
      Log(`Could not obtain parameters (key= <${key}>, token= <${refreshToken}>)!`);
    }
  }
  
  return accessToken;
};


/**
 * GetRefreshTokenSchwab()
 *
 * Obtain a valid refresh token for Schwab API queries
 *
 */
function GetRefreshTokenSchwab(sheetID, verbose)
{
  var refreshTokenExpirationTime= null;
  const refreshTokenTTLOffsetDays= 7;
  const refreshTokenStaleLoggingThrottle= 60 * 60 * 24;
  const currentTime= new Date();
  const refreshTokenCopy= GetValueByName(sheetID, "ParameterSchwabTokenRefreshSaved", verbose);
  var refreshToken= GetValueByName(sheetID, "ParameterSchwabTokenRefresh", verbose);

  if (!refreshToken)
  {
    // Missing refresh token
    LogThrottled(sheetID, `Refresh token missing <${refreshToken}>-- obtain a new one!!!`, verbose, refreshTokenStaleLoggingThrottle);
  }
  else if (refreshToken != refreshTokenCopy)
  {
    // Looks like we have a new refresh token -- save a copy and update expiration time
    SetValueByName(sheetID, "ParameterSchwabTokenRefreshSaved", refreshToken, verbose);

    refreshTokenExpirationTime= currentTime;
    refreshTokenExpirationTime.setDate(refreshTokenExpirationTime.getDate() + refreshTokenTTLOffsetDays);
    SetValueByName(sheetID, "ParameterSchwabTokenRefreshTimeStamp", refreshTokenExpirationTime, verbose);
  }
  else
  {
    // The refresh token has not changed -- get its expiration time stamp
    refreshTokenExpirationTime= GetValueByName(sheetID, "ParameterSchwabTokenRefreshTimeStamp", verbose);
  }
  
  LogVerbose(`Refresh token: ${refreshToken}`, verbose);
  LogVerbose(`Refresh token expiration: ${refreshTokenExpirationTime}`, verbose);
    
  if (currentTime > refreshTokenExpirationTime)
  {
    // Current refresh token has also expired or has no expiration value -- report and invalidate
    LogThrottled(sheetID, "Refresh token has gone stale -- obtain a new one!!!", verbose, refreshTokenStaleLoggingThrottle);

    refreshToken= null;
  }

  return refreshToken;
};


/**
 * GetURLSchwab()
 *
 * Fetch the supplied Schwab API URL via GET
 */
function GetURLSchwab(sheetID, url, verbose)
{
  const headers= ComposeHeadersSchwab(sheetID, verbose);
  var response= null;
  var content= null;
  
  if (headers)
  {
    response= FetchURLSchwab(sheetID, url, headers, "get", null, verbose);
  }
  else
  {
    // Missing parameters
    LogThrottled(sheetID, `Missing parameters: headers= ${headers}`);
  }

  if (response)
  {
    content= ExtractContentSchwab(response);
  }
  else
  {
    // Missing response
    LogThrottled(sheetID, `Received no response for query  <${url}>`);
  }

  return content;
};


/**
 * PostURLSchwab()
 *
 * Fetch the supplied Schwab API URL via POST
 */
function PostURLSchwab(url, headers, formData, verbose)
{

};


/**
 * FetchURLSchwab()
 *
 * Fetch the supplied Schwab API URL
 */
function FetchURLSchwab(sheetID, url, headers, method, payload, verbose)
{
  var response= null;
  var options= {};

  if (headers)
  {
    options["headers"]= headers;
  }

  if (method)
  {
    options["method"]= method;
  }
  else
  {
    options["method"]= "get";
  }

  if (payload)
  {
    options["payload"]= payload;
  }
  
  try
  {
    response= UrlFetchApp.fetch(url, options);
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
function ExtractContentSchwab(response)
{
  // Declare constants and local variables
  const responseOK= 200;
  var contentParsed= null;
  
  if (response.getResponseCode() == responseOK)
  {
    // Looks like we have a valid data response
    const content= response.getContentText();
    contentParsed= JSON.parse(content);
  }
  else
  {
    Log(`Data query returned error code <${response.getResponseCode()}>.`);
    Log(response.getAllHeaders());
  }

  return contentParsed;
};


/**
 * ExtractPriceSchwab()
 *
 * Extract price from a given list of contracts for a specific strike and expiration
 *
 */
function ExtractPriceSchwab(quotes, amount, preferWeekly)
{
  // Declare constants and local variables
  const labelSymbol= "symbol";
  const labelLastPrice= "last";
  const labelClosePrice= "closePrice";
  const labelBidPrice= "bid";
  const labelAskPrice= "ask";
  const labelDelta= "delta";
  const deltaBad= "-999.0";
  const symbolDelimiter= " ";
  const weekly= "W";
  var price= null;

  if (preferWeekly == undefined)
  {
    // Usually, prefer weekly expirations
    preferWeekly= true;
  }

  for (quote of quotes)
  {
    if (quote[labelDelta] != deltaBad)
    {
      const last= parseFloat(quote[labelLastPrice]);
      const close= parseFloat(quote[labelClosePrice]);
      const bid= parseFloat(quote[labelBidPrice]);
      const ask= parseFloat(quote[labelAskPrice]);

      if (bid > 0 && ask > 0)
      {
        // Make sure we hve some semblance of liquidity
        if (last > 0 && last > bid && last < ask)
        {
          // Last price seems valid
          price= {price : last, contract : quote[labelSymbol]};
        }
        else if (close > 0 && close > bid && close < ask)
        {
          // Close price seems valid
          price= {price : close, contract : quote[labelSymbol]};
        }
        else
        {
          // Use the mid point of the spread for price
          price= {price : bid + (ask - bid) / 2, contract : quote[labelSymbol]};
        }
      }

      if (price > amount)
      {
        // Avoid extreme strikes
        price= null;
      }

      if (preferWeekly == quote[labelSymbol].split(symbolDelimiter)[0].endsWith(weekly))
      {
        // Found the preferred type
        break;
      }
    }
  }

  return price;
};