/**
 * GetQuotesTradierAll()
 *
 * Obtain option prices from Tradier
 *
 */
function GetQuotesTradierAll(symbols, labels, urlHead, verbose)
{
  // Declare constants and local variables
  var firstDataRow= 1;
  var symbolColumn= 0;
  var labelURL= "URL";
  var url= null;
  var urls= {};
  var response= null;
  var prices= {};
  
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
  
  // Convert unique symbol keys into a plain array and fetch them with one request
  url= ConstructUrlQuoteTradier(Object.keys(urls), verbose);
  if (url)
  {
    try
    {
      //Logger.log("[GetQuotesTradierAll] Composed URL:");
      //Logger.log(url);
      
      response= UrlFetchApp.fetch(url, ComposeHeadersTradier(verbose));
    }
    catch (error)
    {
      return "[GetQuotesTradierAll] ".concat(error);
    }
  
    if (response)
    {
      // Data fetched -- extract
      prices= ExtractDataTradier(response, urls, labels, verbose);
    }
    else
    {
      // Failed to fetch web pages
      return "[GetQuotesTradierAll] Could not fetch quotes!";
    }
  }
  else
  {
    // No prices to fetch?
    return "[GetQuotesTradierAll] Could not compile queries.";
  }
  
  return prices;
};


/**
 * ExtractDataTradier()
 *
 * Extract pricing data from Tradier JSON result
 *
 */
function ExtractDataTradier(response, urls, labels, verbose)
{
  // Declare constants and local variables
  var labelQuotes= "quotes";
  var labelQuote= "quote";
  var labelSymbol= "symbol";
  var labelGreeks= "greeks";
  var labelURL= "URL";
  var labelDebug= "Debug";
  var responseOK= 200;
  var prices= {};
  
  //labelDebug= undefined;
  
  if (response.getResponseCode() == responseOK)
  {
    // Looks like we have a valid data response
    var content= response.getContentText();
    var quotes= JSON.parse(content)[labelQuotes][labelQuote];
    
    //Logger.log("[ExtractDataTradier] Headers:");
    //Logger.log(response.getAllHeaders());
    
    //Logger.log("[ExtractDataTradier] Content:");
    //Logger.log(content);
    
    if (quotes)
    {
      // Looks like we have valid quotes
      if (quotes[labelSymbol] != undefined)
      {
        // If we got a single response, convert it to an array (*sigh*)
        quotes= [quotes];
      }
      
      for (var quote in quotes)
      {
        // Process each returned quote
        if (quotes[quote][labelSymbol] != undefined)
        {
          // Symbol exists
          prices[quotes[quote][labelSymbol]]= {};
          prices[quotes[quote][labelSymbol]][labelURL]= urls[quotes[quote][labelSymbol]];
          
          if (labelDebug != undefined)
          {
            // debug activated -- commit raw data
            prices[quotes[quote][labelSymbol]][labelDebug]= quotes[quote];
          }
          
          for (var label in labels)
          {
            // Pull a value for each desired label
            if (quotes[quote][labels[label].toLowerCase()] == undefined)
            {
              // No top-level data -- check for Greeks
              if (quotes[quote][labelGreeks] == undefined)
              {
                prices[quotes[quote][labelSymbol]][labels[label]]= "no data (no greeks)";
              }
              else if (quotes[quote][labelGreeks][labels[label].toLowerCase()] == undefined)
              {
                prices[quotes[quote][labelSymbol]][labels[label]]= "no data";
              }
              else
              {
                prices[quotes[quote][labelSymbol]][labels[label]]= quotes[quote][labelGreeks][labels[label].toLowerCase()];
              }
            }
            else
            {
              prices[quotes[quote][labelSymbol]][labels[label]]= quotes[quote][labels[label].toLowerCase()];
            }
          }
        }
      }
    }
    else
    {
      Logger.log("[ExtractDataTradier] Data query returned no content");
      Logger.log(content);
    }
  }
  else
  {
    Logger.log("[ExtractDataTradier] Data query returned error code <%s>.", response.getResponseCode());
    Logger.log(response.getAllHeaders());
  }
  
  return prices;
};


/**
 * ConstructUrlQuoteTradier()
 *
 * Construct a Tradier quote query URL for specified symbols
 *
 */
function ConstructUrlQuoteTradier(symbols, verbose)
{
  // Declare constants and local variables
  var urlHead= "https://sandbox.tradier.com/v1/markets/quotes?symbols=";
  //var urlHead= "https://api.tradier.com/v1/markets/quotes?symbols=";
  var urlTail= "&greeks=true";
  
  //urlTail= "&greeks=false";
  
  //return urlHead + symbols.slice(0, 8).join(",") + urlTail;
  
  return urlHead + symbols.join(",") + urlTail;
};


/**
 * ComposeHeadersTradier()
 *
 * Compose required headers for our requests
 */
function ComposeHeadersTradier(verbose)
{
  // declare constants and local variables
  var headers= { 'Accept' : 'application/json', 'Authorization' : 'Bearer mBLunvxV0ZjUoNe6445c874YR4HN'};
  
  return { 'headers' : headers };
};


/**
 * GetExpirationsForSymbolTradier()
 *
 * Obtain a list of contract expirations for a given underlying
 *
 */
function GetExpirationsForSymbolTradier(symbol, verbose)
{
  // Declare constants and local variables
  var url= null;
  var response= null;
  var expirations= null;
  
  url= ConstructUrlExpirationsTradier(symbol, verbose);
  if (url)
  {
    try
    {
      response= UrlFetchApp.fetch(url, ComposeHeadersTradier(verbose));
    }
    catch (error)
    {
      return "[GetExpirationsForSymbolTradier] ".concat(error);
    }
  
    if (response)
    {
      // Data fetched -- extract
      expirations= ExtractExpirationDatesTradier(response, verbose);
    }
    else
    {
      // Failed to fetch results
      return "[GetExpirationsForSymbolTradier] Could not fetch expirations!";
    }
  }
  else
  {
    // No prices to fetch?
    return "[GetExpirationsForSymbolTradier] Could not compile query.";
  }
  
  return expirations;
};


/**
 * ConstructUrlExpirationsTradier()
 *
 * Construct a Tradier expirations query URL for specified symbols
 *
 */
function ConstructUrlExpirationsTradier(symbol, verbose)
{
  // Declare constants and local variables
  var urlHead= "https://sandbox.tradier.com/v1/markets/options/expirations?symbol=";
  
  return urlHead + symbol;
};


/**
 * ExtractExpirationDatesTradier()
 *
 * Extract expiration dates from Tradier JSON result
 *
 */
function ExtractExpirationDatesTradier(response, verbose)
{
  // Declare constants and local variables
  var labelExpirations= "expirations";
  var labelDate= "date";
  var expirations= null;
  var responseOK= 200;
  
  if (response.getResponseCode() == responseOK)
  {
    // Looks like we have a valid data response
    var content= response.getContentText();
    expirations= JSON.parse(content)[labelExpirations][labelDate];
    
    if (!expirations)
    {
      Logger.log("[ExtractExpirationsTradier] Data query returned no content");
      Logger.log(content);
    }
  }
  else
  {
    Logger.log("[ExtractExpirationsTradier] Data query returned error code <%s>.", response.getResponseCode());
    Logger.log(response.getAllHeaders());
  }
  
  return expirations;
};