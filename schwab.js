/**
 * GetQuotesSchwab()
 *
 * Obtain prices from Schwab
 *
 */
function GetQuotesSchwab(id, symbols, labels, urlHead, verbose)
{
  // Declare constants and local variables
  var firstDataRow= 1;
  var symbolColumn= 0;
  var url= null;
  var urls= {};
  var symbolMap= {};
  var response= null;
  var prices= {};
  
  var headers= ComposeHeadersSchwab(id, verbose);
  
  if (symbols && headers)
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
    
    // Option quotes from Schwab require symbol remapping
    symbolMap= RemapSymbolsSchwab(id, Object.keys(urls), verbose);
    url= ConstructUrlQuoteSchwab(Object.keys(symbolMap));

    if (url)
    {
      try
      {
        response= UrlFetchApp.fetch(url, {'headers' : headers});
      }
      catch (error)
      {
        return "[GetQuotesSchwab] ".concat(error);
      }
      
      if (response)
      {
        // Data fetched -- extract
        prices= ExtractQuotesSchwab(response, symbolMap, urls, labels, verbose);
      }
      else
      {
        // Failed to fetch web pages
        return "[GetQuotesSchwab] Could not fetch quotes!";
      }
    }
    else
    {
      // No prices to fetch?
      return "[GetQuotesSchwab] Could not compile queries.";
    }
  }
  else
  {
    // Missing parameters
    Logger.log("[GetQuotesSchwab] Missing parameters: apiKey= %s, symbols= %s, headers= %s", apiKey, symbols, headers);
  }
  
  return prices;
};


/**
 * ExtractQuotesSchwab()
 *
 * Extract pricing data from Schwab JSON result
 *
 */
function ExtractQuotesSchwab(response, symbolMap, urls, labels, verbose)
{
  // Declare constants and local variables
  var quotes= ExtractContentSchwab(response);
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
    
    // Lastly, define type exeptions
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
    Logger.log("[ExtractQuotesSchwab] Could not obtain quotes!");
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
function RemapSymbolsSchwab(id, symbols, verbose)
{
  // Declare constants and local variables
  const optionDetailsLength= "YYMMDDT00000000".length;
  const symbolMapProvided= GetParameters(id, "ParameterMapSymbols", verbose);
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

      // Temporary compatibility for the TDA to Schwab transition
      // var mappedSymbol= symbolMapProvided[quoteSymbol.toUpperCase()];
      // if (mappedSymbol.endsWith(".X"))
      // {
      //   mappedSymbol= mappedSymbol.slice(0, -2);
      // }
      // else if (mappedSymbol == "BRK.B")
      // {
      //   mappedSymbol= "BRK/B";
      // }
      // symbolMap[mappedSymbol]= quoteSymbol;


      // Original: return after compatibility testing
      // symbolMap[symbolMapProvided[quoteSymbol.toUpperCase()]]= quoteSymbol;
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
  var contracts= {};
  var results= [];
  
  for (var index in symbols)
  {
    // Find candidates for each requested symbol
    symbol= symbols[index][0];
    
    if (symbol)
    {
      contracts= GetContractsForSymbolByExpirationSchwab(sheetID, symbol, dte, labelPuts, labelCalls, verbose);
    
      if (typeof contracts == "string")
      {
        // Looks like we have an error message
        Logger.log("[GetIndexStrangleContractsSchwab] Could not get contracts! (%s)", contracts);
      }
      else if (contracts)
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
        Logger.log("[GetIndexStrangleContractsSchwab] Failed to obtain a list of valid option contracts <%s>!", contracts);
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
    Logger.log("[GetContractsForSymbolByExpirationSchwab] Could not fetch expirations!");
  }

  return contracts;
};


/**
 * GetChainForSymbolByExpirationSchwab()
 *
 * Obtain option contract expiration dates for a given symbol
 *
 */
function GetChainForSymbolByExpirationSchwab(sheetID, symbol, dteEarliest, dteLatest, verbose)
{
  // Declare constants and local variables
  var apiKey= GetValueByName(sheetID, "ParameterSchwabKey", verbose);
  var headers= ComposeHeadersSchwab(sheetID, verbose);
  var url= null;
  var response= null;
  
  if (apiKey && headers)
  {
    // We have viable headers and query parameters, proceed
    url= ConstructUrlExpirationsSchwab(apiKey, symbol, dteEarliest, dteLatest);

    if (verbose)
    {
      Logger.log("[GetContractsForSymbolByExpirationSchwab] Query: %s", url);
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
        return "[GetChainForSymbolByExpirationSchwab] ".concat(error);
      }
    }
    else
    {
      // No prices to fetch?
      Logger.log("[GetContractsForSymbolByExpirationSchwab] Could not compile query.");
    }
  }
  else
  {
    // Missing parameters
    Logger.log("[GetChainForSymbolByExpirationSchwab] Missing parameters: apiKey= %s, headers= %s", apiKey, headers);
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
        Logger.log("[ExtractExpirationDatesSchwab] Data query returned no content");
        Logger.log(content);
      }
    }
  }
  else
  {
    Logger.log("[ExtractExpirationDatesSchwab] Query returned no data!");
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
      Logger.log("[ExtractExpirationDatesSchwab] Data query returned no viable expirations: <%s>", expirations);
      Logger.log(contentParsed);
    }
  }
  else
  {
    Logger.log("[ExtractExpirationDatesSchwab] Query returned no data!");
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
  const urlSymbols= "symbols=" + symbols.join(",");
  const urlFields= "&fields=quote";
  const urlIndicative= "&indicative=false";
  
  return urlHead + urlSymbols + urlFields + urlIndicative;
};


/**
 * ConstructUrlExpirationsSchwab()
 *
 * Construct a URL to obtain a list of contracts with expiration dates from Schwab
 */
function ConstructUrlExpirationsSchwab(apiKey, underlying, dteEarliest, dteLatest)
{
  // Declare constants and local variables
  var urlHead= "https://api.schwabapi.com/marketdata/v1/chains?strikeCount=500&strategy=SINGLE";
  var urlSymbol= "&symbol=" + underlying;
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
  
  return urlHead + urlSymbol + urlFromDate + urlToDate;
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
  var headers= {};
  
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
    Logger.log("[ComposeHeadersSchwab] Cannot request quotes without a valid access token [%s]!", accessToken);
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
  var accessToken= null;
  
  // A 5-minute buffer
  const accessTokenTTLOffset= 300;
  
  // First, validate preserved access token
  const accessTokenExpiration= GetValueByName(sheetID, "ParameterSchwabTokenAccessExpiration", verbose);
  if (accessTokenExpiration && (currentTime < accessTokenExpiration))
  {
    // our preserved access token remains valid
    accessToken= GetValueByName(sheetID, "ParameterSchwabTokenAccess", verbose);
  }
  
  if (!accessToken)
  {
    // our preserved acces token has expired or is otherwise invalid -- refresh it
    const key= GetValueByName(sheetID, "ParameterSchwabKey", verbose);
    var refreshToken= GetValueByName(sheetID, "ParameterSchwabTokenRefresh", verbose);
    var refreshTokenExpiration= GetValueByName(sheetID, "ParameterSchwabTokenRefreshExpiration", verbose);
    var response= null;

    if (verbose)
    {
      Logger.log("[GetAccessTokenSchwab] Refreshing invalid access token (expiration stamp <%s>)...", accessTokenExpiration);
      Logger.log("[GetAccessTokenSchwab] Old access token: %s", accessToken);
      Logger.log("[GetAccessTokenSchwab] Refresh token expiration stamp <%s>...", refreshTokenExpiration);
      Logger.log("[GetAccessTokenSchwab] Refresh token: %s", refreshToken);
    }
    
    if (!refreshTokenExpiration || (currentTime > refreshTokenExpiration))
    {
      // Preserved refresh token has also expired or has no expiration value
      if (verbose)
      {
        Logger.log("[GetAccessTokenSchwab] Refresh token has gone stale -- obtain a new one!!!");
      }
    }
    else
    {
      // Request a new access token using our valid refreh token
      if (verbose)
      {
        Logger.log("[GetAccessTokenSchwab] Refresh token remains valid...");
      }

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
            // const keyRefreshToken= "refresh_token";
            // const keyRefreshTokenTTL= "refresh_token_expires_in";
            const contentParsed= ExtractContentSchwab(response);
            
            const accessTokenTTL= contentParsed[keyAccessTokenTTL];
            // const refreshToken= contentParsed[keyRefreshToken];
            // const refreshTokenTTL= contentParsed[keyRefreshTokenTTL];
            var expiration= null;

            accessToken= contentParsed[keyAccessToken];
            
            // if (refreshToken && refreshTokenTTL)
            // {
            //   // Preserve the new Refresh Token and its expiration time
            //   expiration= new Date();
            //   expiration.setSeconds(expiration.getSeconds() + refreshTokenTTL);
            //   expiration.setDate(expiration.getDate() - 14);
              
            //   SetValueByName(sheetID, "ParameterSchwabTokenRefresh", refreshToken, verbose);
            //   SetValueByName(sheetID, "ParameterSchwabTokenRefreshExpiration", expiration, verbose);
            // }
            
            if (accessToken && accessTokenTTL)
            {
              // Preserve the new Access Token and its expiration time
              expiration= new Date();
              expiration.setSeconds(expiration.getSeconds() + accessTokenTTL - accessTokenTTLOffset);
              
              SetValueByName(sheetID, "ParameterSchwabTokenAccess", accessToken, verbose);
              SetValueByName(sheetID, "ParameterSchwabTokenAccessExpiration", expiration, verbose);
            }
            else
            {
              Logger.log("[GetAccessTokenSchwab] Failed to obtain refreshed access token [%s] and its time-to-live [%s]!", accessToken, accessTokenTTL);
            }
          }
          else
          {
            Logger.log("[GetAccessTokenSchwab] Error repsonse code: [%s]", response.getResponseCode());
          }
        }
        catch (error)
        {
          return "[GetAccessTokenSchwab] ".concat(error);
        }
      }
      else if (verbose)
      {
        Logger.log("[GetAccessTokenSchwab] Could not obtain parameters (key= <%s>, token= <%s>)!", key, refreshToken);
      }
    }
  }
  
  return accessToken;
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
    Logger.log("[ExtractContentSchwab] Data query returned error code <%s>.", response.getResponseCode());
    Logger.log(response.getAllHeaders());
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