/**
 * GetQuotesTDA()
 *
 * Obtain prices from TDA
 *
 */
function GetQuotesTDA(id, symbols, labels, urlHead, verbose)
{
  // Declare constants and local variables
  var firstDataRow= 1;
  var symbolColumn= 0;
  var url= null;
  var urls= {};
  var symbolMap= {};
  var response= null;
  var prices= {};
  
  var apiKey= GetValueByName(id, "ParameterTDAKey", verbose);
  var headers= ComposeHeadersTDA(id, verbose);
  
  if (symbols && apiKey && headers)
  {
    // We have viable headers and query parameters, proceed
    
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
    
    // Option quotes from TDA require symbol remapping
    symbolMap= RemapSymbolsTDA(id, Object.keys(urls), verbose);
    url= ConstructUrlQuoteTDA(apiKey, Object.keys(symbolMap));

    if (url)
    {
      try
      {
        response= UrlFetchApp.fetch(url, {'headers' : headers});
      }
      catch (error)
      {
        return "[GetQuotesTDA] ".concat(error);
      }
      
      if (response)
      {
        // Data fetched -- extract
        prices= ExtractQuotesTDA(response, symbolMap, urls, labels, verbose);
      }
      else
      {
        // Failed to fetch web pages
        return "[GetQuotesTDA] Could not fetch quotes!";
      }
    }
    else
    {
      // No prices to fetch?
      return "[GetQuotesTDA] Could not compile queries.";
    }
  }
  else
  {
    // Missing parameters
    Logger.log("[GetQuotesTDA] Missing parameters: apiKey= %s, symbols= %s, headers= %s", apiKey, symbols, headers);
  }
  
  return prices;
};


/**
 * ExtractQuotesTDA()
 *
 * Extract pricing data from Tradier JSON result
 *
 */
function ExtractQuotesTDA(response, symbolMap, urls, labels, verbose)
{
  // Declare constants and local variables
  var responseOK= 200;
  var prices= {};
  
  if (response.getResponseCode() == responseOK)
  {
    // Looks like we have a valid data response
    var content= response.getContentText();
    var quotes= JSON.parse(content);
    
    if (quotes)
    {
      // We have quotes -- process them
      
      // First, define interesting quote parameters
      const labelType= "assetType";
      const labelSymbol= "symbol";
      const labelChange= "netChange";
      const labelChangeInDouble= "changeInDouble";
      const labelNetChange= "netChange";
      const labelLast= "lastPrice";
      const labelLastInDouble= "lastPriceInDouble";
      const labelNAV= "nAV";
      const labelClose= "closePrice";
      const labelCloseInDouble= "closePriceInDouble";
      const labelClosePrice= "closePrice";
      const labelBid= "bidPrice";
      const labelAsk= "askPrice";
      const labelDelta= "delta";
      const labelURL= "URL";
      const labelDebug= "Debug";
      
      // Second, seed the default table column to quote parameter map
      const columnBid= "Bid";
      const columnAsk= "Ask";
      const columnLast= "Last";
      const columnClose= "Close";
      const columnChange= "Change";
      const columnDelta= "Delta";
      var labelMap= {};
      labelMap[columnBid]= labelBid;
      labelMap[columnAsk]= labelAsk;
      labelMap[columnLast]= labelLast;
      labelMap[columnClose]= labelClose;
      labelMap[columnChange]= labelChange;
      labelMap[columnDelta]= labelDelta;
      
      // Lastly, define type exeptions
      const typeMutualFund= "MUTUAL_FUND";
      const typeFuture= "FUTURE";
      
      for (const quote in quotes)
      {
        // Process each returned quote
        if (quotes[quote][labelSymbol] != undefined)
        {
          // Symbol exists
          prices[symbolMap[quotes[quote][labelSymbol]]]= {};
          prices[symbolMap[quotes[quote][labelSymbol]]][labelURL]= urls[symbolMap[quotes[quote][labelSymbol]]];
          
          if (labelDebug != undefined)
          {
            // debug activated -- commit raw data
            prices[symbolMap[quotes[quote][labelSymbol]]][labelDebug]= quotes[quote];
          }
          
          // Adjust labelMap based quote specifics
          if (quotes[quote][labelType] == typeMutualFund)
          {
            // Use NAV as last price for mutual funds
            labelMap[columnLast]= labelNAV;
            labelMap[columnChange]= labelNetChange;
            labelMap[columnClose]= labelClosePrice;
          }
          else if (quotes[quote][labelType] == typeFuture)
          {
            // Use NAV as last price for mutual funds
            labelMap[columnLast]= labelLastInDouble;
            labelMap[columnChange]= labelChangeInDouble;
            labelMap[columnClose]= labelCloseInDouble;
          }
          else
          {
            // Revert to defaults
            labelMap[columnLast]= labelLast;
            labelMap[columnChange]= labelChange;
            labelMap[columnClose]= labelClose;
          }
          
          for (const label in labels)
          {
            // Pull a value for each desired label
            if (quotes[quote][labelMap[labels[label]]] == undefined)
            {
              // No top-level data
              prices[symbolMap[quotes[quote][labelSymbol]]][labels[label]]= "no data";
            }
            else
            {
              prices[symbolMap[quotes[quote][labelSymbol]]][labels[label]]= quotes[quote][labelMap[labels[label]]];
            }
          }
        }
      }
    }
    else
    {
      Logger.log("[ExtractQuotesTDA] Data query returned no content");
      Logger.log(content);
    }
  }
  else
  {
    Logger.log("[ExtractQuotesTDA] Data query returned error code <%s>.", response.getResponseCode());
    Logger.log(response.getAllHeaders());
  }
  
  return prices;
};


/**
 * RemapSymbolsTDA()
 *
 * Remap symbols to the TDA format
 *   covers: options, indexes, futures, special cases (e.g., US10Y)
 *
 */
function RemapSymbolsTDA(id, symbols, verbose)
{
  // Declare constants and local variables
  var symbolMap= {};
  var symbol= "";
  
  var symbolTDA= "";
  var underlying= "";
  var date= "";
  var month= "";
  var year= "";
  var type= "";
  var strike= "";
  var symbolIndex= 0;
  var optionDetailsLength= "YYMMDDT00000000".length;
  var symbolMapProvided= GetParameters(id, "ParameterMapSymbols", verbose);
  
  for (const symbol of symbols)
  {
    if (symbolMapProvided[symbol] != undefined)
    {
      symbolMap[symbolMapProvided[symbol.toUpperCase()]]= symbol;
    }
    else if (symbol.length > optionDetailsLength)
    {
      // Re-map each apparent option symbol
    
      underlying= symbol.slice(0, -optionDetailsLength);
      
      symbolIndex= underlying.length;
      year= symbol.slice(symbolIndex, symbolIndex+= 2);
      month= symbol.slice(symbolIndex, symbolIndex+= 2);
      date= symbol.slice(symbolIndex, symbolIndex+= 2);
      
      type= symbol.slice(symbolIndex, symbolIndex+= 1);
      
      strike= symbol.slice(symbolIndex);
      strike= strike / 1000;
      
      symbolTDA= underlying + "_" + month + date + year + type + strike;
      symbolMap[symbolTDA]= symbol;
    }
    else
    {
      // Preserve all other symbols
      
      symbolMap[symbol]= symbol;
    }
  }
  
  return symbolMap;
};


/**
 * GetIndexStrangleContractsTDA()
 *
 * Obtain and select best matching strangles
 *
 */
function GetIndexStrangleContractsTDA(sheetID, symbols, dte, deltaCall, deltaPut, verbose)
{
  // Declare constants and local variables
  const labelPuts= "putExpDateMap";
  const labelCalls= "callExpDateMap";
  var symbol= null;
  var contracts= {};
  var results= [];
  
  for (var index in symbols)
  {
    // Find candidates for each requested symbol
    symbol= symbols[index][0];
    
    if (symbol)
    {
      contracts= GetContractsForSymbolByExpirationTDA(sheetID, symbol, dte, labelPuts, labelCalls, verbose);
    
      if (typeof contracts == "string")
      {
        // Looks like we have an error message
        Logger.log("[GetIndexStrangleContractsTDA] Could not get contracts! (%s)", contracts);
      }
      else if (contracts)
      {
        // We have valid contracts -- sift through them for best delta matches
        if (contracts[labelPuts])
        {
          // Viable puts data
          results= results.concat(FindBestDeltaMatchTDA(contracts[labelPuts], deltaPut, verbose));
        }
        
        if (contracts[labelCalls])
        {
          // Viable calls data
          results= results.concat(FindBestDeltaMatchTDA(contracts[labelCalls], deltaCall, verbose));
        }
      }
      else
      {
        Logger.log("[GetIndexStrangleContractsTDA] Failed to obtain a list of valid option contracts <%s>!", contracts);
      }
    }
  }
  
  return results;
};


/**
 * FindBestDeltaMatchTDA()
 *
 * Find a contract in the given set with closest delta value match
 *
 */
function FindBestDeltaMatchTDA(contracts, deltaTarget, verbose)
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
  const symbolDelimiter= "_";
  const weekly= "W";
  
  deltaTarget/= 100;
  const deltaTargetMinimum= deltaTarget - 0.01;
  const deltaTargetMaximum= deltaTarget + 0.01;
  
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
 * GetContractsForSymbolByExpirationTDA()
 *
 * Obtain option contract expiration dates for a given symbol
 *
 */
function GetContractsForSymbolByExpirationTDA(sheetID, symbol, dte, labelPuts, labelCalls, verbose)
{
  // Declare constants and local variables
  var contracts= {};
  const response= GetChainForSymbolByExpirationTDA(sheetID, symbol, dte, dte, verbose);
  
  if (response)
  {
    // Data fetched -- extract
    contracts= ExtractEarliestContractsTDA(response, labelPuts, labelCalls, verbose);
  }
  else
  {
    // Failed to fetch results
    Logger.log("[GetContractsForSymbolByExpirationTDA] Could not fetch expirations!");
  }

  return contracts;
};


/**
 * GetChainForSymbolByExpirationTDA()
 *
 * Obtain option contract expiration dates for a given symbol
 *
 */
function GetChainForSymbolByExpirationTDA(sheetID, symbol, dteEarliest, dteLatest, verbose)
{
  // Declare constants and local variables
  var apiKey= GetValueByName(sheetID, "ParameterTDAKey", verbose);
  var headers= ComposeHeadersTDA(sheetID, verbose);
  var url= null;
  var response= null;
  
  if (apiKey && headers)
  {
    // We have viable headers and query parameters, proceed
    url= ConstructUrlExpirationsTDA(apiKey, symbol, dteEarliest, dteLatest);

    if (verbose)
    {
      Logger.log("[GetContractsForSymbolByExpirationTDA] Query: %s", url);
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
        return "[GetChainForSymbolByExpirationTDA] ".concat(error);
      }
    }
    else
    {
      // No prices to fetch?
      Logger.log("[GetContractsForSymbolByExpirationTDA] Could not compile query.");
    }
  }
  else
  {
    // Missing parameters
    Logger.log("[GetChainForSymbolByExpirationTDA] Missing parameters: apiKey= %s, headers= %s", apiKey, headers);
  }

  return response;
}


/**
 * ExtractEarliestContractsTDA()
 *
 * Extract contract expiration dates from a list of returned contracts
 */
function ExtractEarliestContractsTDA(response, labelPuts, labelCalls, verbose)
{
  // Declare constants and local variables
  var content= null;
  var contentParsed= null;
  var labelDate= null;
  var expirations= [];
  var contractTypes= [labelPuts, labelCalls];
  var contractType= null;
  var contracts= {};
  const responseOK= 200;
  
  if (response.getResponseCode() == responseOK)
  {
    // Looks like we have a valid data response
    content= response.getContentText();
    contentParsed= JSON.parse(content);
    
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
        Logger.log("[ExtractExpirationDatesTDA] Data query returned no content");
        Logger.log(content);
      }
    }
  }
  else
  {
    Logger.log("[ExtractExpirationDatesTDA] Data query returned error code <%s>.", response.getResponseCode());
    Logger.log(response.getAllHeaders());
  }
  
  return contracts;
};


/**
 * ExtractExpirationsTDA()
 *
 * Extract matching contract expiration dates from an option chain
 */
function ExtractExpirationsTDA(response, expirationTargets, verbose)
{
  // Declare constants and local variables
  const responseOK= 200;
  const labelPuts= "putExpDateMap";
  
  if (response.getResponseCode() == responseOK)
  {
    // Looks like we have a valid data response
    const content= response.getContentText();
    const contentParsed= JSON.parse(content);
    var expirations= Object.keys(contentParsed[labelPuts]).sort();
    var expirationsMapped= [];
    var expirationDate= null;
    var expirationTarget= null;

    if (expirations)
    {
      // We have a list of expirations -- now map them to desired DTE targets
      for (var target in expirationTargets)
      {
        // Match an expiration date to each valid target
        if (typeof expirationTargets[target][0] == "number")
        {
          expirationTarget= new Date();
          expirationTarget.setDate(expirationTarget.getDate() + expirationTargets[target][0]);

          for (var expiration in expirations)
          {
            // Find the earliest expiration date which satisfies our days-to-expiration constraint
            expirationDate= new Date(expirations[expiration].split(":")[0]);

            if (expirationDate >= expirationTarget)
            {
              // We found the earliest expiration date
              const labelSymbol= "symbol";
              const symbolDelimiter= "_";
              const weekly= "W";
              const strike= Object.keys(contentParsed[labelPuts][expirations[expiration]])[0];
              var underlying= "";

              // Determine the udnerlying symbol from a given quote
              for (var quote in contentParsed[labelPuts][expirations[expiration]][strike])
              {
                underlying= contentParsed[labelPuts][expirations[expiration]][strike][quote][labelSymbol].split(symbolDelimiter)[0];
                if (underlying.endsWith(weekly))
                {
                  // Prefer weeklies
                  break;
                }
              }

              expirationsMapped[target]= [expirations[expiration].split(":")[0], underlying];
              break;
            }
          }
        }
      }
    }
    else
    {
      Logger.log("[ExtractExpirationDatesTDA] Data query returned no viable expirations: <%s>", expirations);
      Logger.log(contentParsed);
    }
  }
  else
  {
    Logger.log("[ExtractExpirationDatesTDA] Data query returned error code <%s>.", response.getResponseCode());
    Logger.log(response.getAllHeaders());
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
 * ConstructUrlQuoteTDA()
 *
 * Construct a TDA quote query URL for specified symbols
 *
 */
function ConstructUrlQuoteTDA(apiKey, symbols)
{
  // Declare constants and local variables
  var urlHead= "https://api.tdameritrade.com/v1/marketdata/quotes?";
  var urlAPIKey= "apikey=" + apiKey;
  var urlSymbols= "&symbol=" + symbols.join(",");
  
  return urlHead + urlAPIKey + urlSymbols;
};


/**
 * ConstructUrlExpirationsTDA()
 *
 * Construct a URL to obtain a list of contracts with expiration dates from TDA
 */
function ConstructUrlExpirationsTDA(apiKey, symbol, dteEarliest, dteLatest)
{
  // Declare constants and local variables
  var urlHead= "https://api.tdameritrade.com/v1/marketdata/chains?strikeCount=500&strategy=SINGLE";
  var urlAPIKey= "&apikey=" + apiKey;
  var urlSymbol= "&symbol=" + symbol;
  var urlFromDate= "&fromDate=";
  var urlToDate= "&toDate=";
  
  // construct dates in yyyy-mm-dd format by abusing the ISO string output method
  var dateParameter= new Date();
  
  // aim for requested DTE
  dateParameter.setDate(dateParameter.getDate() + dteEarliest);
  urlFromDate+= dateParameter.toISOString().split('T').shift();
  
  // collect expirations during the specified window (plus a buffer)
  dateParameter.setDate(dateParameter.getDate() + dteLatest);
  urlToDate+= dateParameter.toISOString().split('T').shift();
  
  return urlHead + urlAPIKey + urlSymbol + urlFromDate + urlToDate;
};


/**
 * ComposeHeadersTDA()
 *
 * Compose required headers for our requests
 */
function ComposeHeadersTDA(sheetID, verbose)
{
  // Declare constants and local variables
  var accessToken= GetAccessTokenTDA(sheetID, verbose);
  var headers= {};
  
  if (accessToken)
  {
    // We have an access token, proceed
    headers= { "Accept" : "application/json", "Authorization" : "Bearer " + accessToken };
  }
  else
  {
    // We did not get an access token -- proceed without one for delayed quotes
    headers= { "Accept" : "application/json" };
    if (verbose)
    {
      Logger.log("[ComposeHeadersTDA] Cannot request real-time quotes without a valid access token [%s]!", accessToken);
    }
  }
  
  return headers;
};


/**
 * GetAccessTokenTDA()
 *
 * Obtain a fresh and valid access token for TDA API queries
 *
 */
function GetAccessTokenTDA(sheetID, verbose)
{
  // Declare constants and local variables
  var currentTime= new Date();
  var accessToken= null;
  
  // A 5-minute buffer
  var accessTokenTTLOffset= 300;
  
  // First, validate preserved access token
  var accessTokenExpiration= GetValueByName(sheetID, "ParameterTDATokenAccessExpiration", verbose);
  if (accessTokenExpiration && (currentTime < accessTokenExpiration))
  {
    // our preserved access token remains valid
    accessToken= GetValueByName(sheetID, "ParameterTDATokenAccess", verbose);
  }
  
  if (!accessToken)
  {
    // our preserved acces token has expired or is otherwise invalid -- refresh it
    var key= GetValueByName(sheetID, "ParameterTDAKey", verbose);
    var refreshToken= GetValueByName(sheetID, "ParameterTDATokenRefresh", verbose);
    var refreshTokenExpiration= GetValueByName(sheetID, "ParameterTDATokenRefreshExpiration", verbose);
    var accessType= "";
    var responseOK= 200;
    var response= null;
    var headers= {};

    if (verbose)
    {
      Logger.log("[GetAccessTokenTDA] Refreshing invalid access token (expiration stamp <%s>)...", accessTokenExpiration);
      Logger.log("[GetAccessTokenTDA] Old access token: %s", accessToken);
      Logger.log("[GetAccessTokenTDA] Refresh token expiration stamp <%s>...", refreshTokenExpiration);
      Logger.log("[GetAccessTokenTDA] Refresh token: %s", refreshToken);
    }
    
    if (!refreshTokenExpiration || (currentTime > refreshTokenExpiration))
    {
      // our preserved refresh token has also expired
      accessType= "offline";
      
      if (verbose)
      {
        Logger.log("[GetAccessTokenTDA] Also refreshing stale refresh token...");
      }
    }
    
    if (key && refreshToken)
    {
      // we have necessary parameters, proceed to obtain token
      var formData= {
        'grant_type': 'refresh_token',
        'refresh_token': refreshToken,
        'access_type': accessType,
        'client_id': key
      };
      
      var options = {
        'method' : 'post',
        'payload' : formData
      };
      
      try
      {
        response= UrlFetchApp.fetch("https://api.tdameritrade.com/v1/oauth2/token", options);
        
        if (response.getResponseCode() == responseOK)
        {
          // Looks like we have a valid data response
          var keyAccessToken= "access_token";
          var keyAccessTokenTTL= "expires_in";
          var keyRefreshToken= "refresh_token";
          var keyRefreshTokenTTL= "refresh_token_expires_in";
          var content= response.getContentText();
          var contentParsed= JSON.parse(content);
          
          accessToken= contentParsed[keyAccessToken];
          var accessTokenTTL= contentParsed[keyAccessTokenTTL];
          var refreshToken= contentParsed[keyRefreshToken];
          var refreshTokenTTL= contentParsed[keyRefreshTokenTTL];
          var expiration= null;
          
          if (refreshToken && refreshTokenTTL)
          {
            // Preserve the new Refresh Token and its expiration time
            expiration= new Date();
            expiration.setSeconds(expiration.getSeconds() + refreshTokenTTL);
            expiration.setDate(expiration.getDate() - 14);
            
            SetValueByName(sheetID, "ParameterTDATokenRefresh", refreshToken, verbose);
            SetValueByName(sheetID, "ParameterTDATokenRefreshExpiration", expiration, verbose);
          }
          
          if (accessToken && accessTokenTTL)
          {
            // Preserve the new Access Token and its expiration time
            expiration= new Date();
            expiration.setSeconds(expiration.getSeconds() + accessTokenTTL - accessTokenTTLOffset);
            
            SetValueByName(sheetID, "ParameterTDATokenAccess", accessToken, verbose);
            SetValueByName(sheetID, "ParameterTDATokenAccessExpiration", expiration, verbose);
          }
          else
          {
            Logger.log("[GetAccessTokenTDA] Failed to obtain refreshed access token [%s] and its time-to-live [%s]!", accessToken, accessTokenTTL);
          }
        }
        else
        {
          Logger.log("[GetAccessTokenTDA] Error repsonse code: [%s]", response.getResponseCode());
        }
      }
      catch (error)
      {
        return "[GetAccessTokenTDA] ".concat(error);
      }
    }
    else if (verbose)
    {
      Logger.log("[GetAccessTokenTDA] Could not obtain parameters (key= <%s>, token= <%s>)!", key, refreshToken);
    }
  }
  
  return accessToken;
};