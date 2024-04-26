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
 * ConstructUrlChainByExpirationSchwab()
 *
 * Construct a URL to obtain a list of contracts for a specific expiration date
 */
function ConstructUrlChainByExpirationSchwab(underlying, expirationDate, contractType, verbose)
{
  // Declare constants and local variables
  const urlHead= "https://api.schwabapi.com/marketdata/v1/chains?";
  const urlSymbol= "symbol=" + underlying;
  const urlCount= "&strikeCount=500";
  const urlStrategy= "&strategy=SINGLE";
  var urlFromDate= "&fromDate=" + expirationDate;
  var urlToDate= "&toDate=" + expirationDate;
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
 * ComposeGetHeadersSchwab()
 *
 * Compose required headers for our requests
 */
function ComposeGetHeadersSchwab(sheetID, verbose)
{
  // Declare constants and local variables
  const accessToken= GetAccessTokenSchwab(sheetID, verbose);
  // const accessToken= GetAccessTokenSchwabLegacy(sheetID, verbose);
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
    const refreshToken = GetRefreshTokenSchwab(sheetID, verbose);

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
          Log(`Failed to obtain refreshed access token [${accessToken}] and its time-to-live [${accessTokenTTL}]!`);
        }
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
  const headers = ComposeGetHeadersSchwab(sheetID, verbose);
  var response = null;
  var content = null;
  
  if (headers)
  {
    response = FetchURLSchwab(sheetID, url, headers, "get", null, verbose);
  }
  else
  {
    // Missing parameters
    LogThrottled(sheetID, `Missing parameters: headers= ${headers}`);
  }

  if (response)
  {
    content = ExtractContentSchwab(response);
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
function PostURLSchwab(sheetID, url, payload, verbose)
{
  const key= GetValueByName(sheetID, "ParameterSchwabKey", verbose);
  var headers = null;
  var response = null;
  var content = null;

  if (key)
  {
    headers = {
      'Authorization': "Basic " + Utilities.base64Encode(key)
    };

    response = FetchURLSchwab(sheetID, url, headers, "post", payload, verbose);
  }
  else
  {
    LogThrottled(sheetID, `Could not read stored key: ${key}`);
  }

  if (response)
  {
    content = ExtractContentSchwab(response);
  }
  else
  {
    // Missing response
    LogThrottled(sheetID, `Received no response for query  <${url}>`);
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