/**
 * GetQuotesNasdaqAll()
 *
 * Obtain option prices from NASDAQ via parallel calls
 *
 */
function GetQuotesNasdaqAll(symbols, labels, options, verbose)
{
  // Declare constants and local variables
  var firstDataRow= 1;
  var optionColumn= 0;
  var labelURL= "URL";
  var url= null;
  var urls= null;
  var requests= {};
  var responses= null;
  var prices= {};
  
  for (var vIndex= firstDataRow; vIndex < symbols.length; vIndex++)
  {
    // Compile a list of unique URLs
    if (symbols[vIndex][optionColumn].length > 0)
    {
      // Live line
       if (options)
       {
         // Configfure for option quotes
         url= ConstructUrlOptionNasdaq(symbols[vIndex][optionColumn], verbose);
       }
      else
      {
        // Configure for regular quotes
        url= ConstructUrlQuoteNasdaq(symbols[vIndex][optionColumn], verbose);
      }
      
      if (url)
      {
        // Create entries for each URL in two mirroring maps for future reconciliation
        requests[url]= symbols[vIndex][optionColumn];
      }
    }
  }
  
  // Convert unique URL keys into a plain array and fetch them in parallel
  if (urls= Object.keys(requests))
  {
    try
    {
      responses= UrlFetchApp.fetchAll(urls);
    }
    catch (error)
    {
      return "[GetQuotesNasdaqAll] ".concat(error);
    }
  
    if (responses)
    {
      // Pages fetched -- extract data
      for (var index in responses)
      {
        // Extract data from each response
        prices[requests[urls[index]]]= ExtractDataNasdaq(responses[index], labels, verbose);
        
        // Append the URL
        prices[requests[urls[index]]][labelURL]= urls[index];
      }
    }
    else
    {
      // Failed to fetch web pages
      return "[GetQuotesNasdaqAll] Could not fetch quotes!";
    }
  }
  else
  {
    // No prices to fetch?
    return "[GetQuotesNasdaqAll] Could not compile queries.";
  }
  
  return prices;
};


/**
 * ExtractDataNasdaq()
 *
 * Extract pricing data from NASDAQ quote pages (HTML source)
 *
 */
function ExtractDataNasdaq(response, labels, verbose)
{
  // Declare constants and local variables
  var firstMatch= 1;
  var responseOK= 200;
  var labelPrefix= "";
  var labelSign= "";
  var values= {};
  var chunks= [];
  var data= null;
  var pattern= null;
  var typePatterns= false;
  
  if (response.getResponseCode() == responseOK)
  {
    // Looks like we have a valid data response
    data= response.getContentText();
    
    if (typePatterns= ExtractTypeNasdaq(data, verbose))
    {
      // We have a known pattern, proceed
      labelPrefix= typePatterns[0];
      
      for (var label in labels)
      {
        if(labelPrefix)
        {
          // Extract price for each type
          pattern= new RegExp(labelPrefix + labels[label].toLowerCase() + "\\D+(>unch</div>)", "i");
          
          if (pattern.test(data))
          {
            // Convert alpha code for unchanged to zero
            values[labels[label]]= 0;
          }
          else
          {
            // Search for the actual value
            pattern= new RegExp(labelPrefix + labels[label].toLowerCase() + "\\D+(\\d+\\.?\\d*)", "i");
          
            chunks= pattern.exec(data);
            if (chunks)
            {
              // We have a match!
              values[labels[label]]= chunks[firstMatch];
              
              if (labelSign= typePatterns[1])
              {
                // Check if it should be negative
                
                pattern= new RegExp(labelPrefix + labels[label].toLowerCase() + "\\D+" + labelSign + "\\D+\\d+\\.?\\d*", "i");
                if (pattern.test(data))
                {
                  values[labels[label]]= -values[labels[label]];
                }
              }
            }
            else
            {
              // No match?!
              values[labels[label]]= "no data";
            }
          }
        }
        else
        {
          // Invalid symbol?
          values[labels[label]]= "invalid symbol";
          
          Logger.log("[ExtractDataNasdaq] Invalid data extraction prefix  <%s>.", labelPrefix);
        }
      }
    }
    else
    {
      // Unknown type pattern?
      for (var label in labels)
      {
        values[labels[label]]= "invalid symbol";
      }
      
      if (verbose)
      {
        Logger.log("[ExtractDataNasdaq] Matched no known page type patterns.");
      }
    }
  }
  else
  {
    // Looks like an error response code
    for (var label in labels)
    {
      values[labels[label]]= "error: " + response.getResponseCode();
    }
    
    Logger.log("[ExtractDataNasdaq] Data query returned error code <%s>.", response.getResponseCode());
    Logger.log(response.getAllHeaders().toSource());
  }
  
  return values;
};


/**
 * ExtractTypeNasdaq()
 *
 * Determine type of data from NASDAQ quote pages (HTML source) and return proper extraction prefix
 *
 */
function ExtractTypeNasdaq(data, verbose)
{
  // Declare constants and local variables
  var patternStock= "Stock Quote";
  var patternFund= "Mutual Fund Price";
  var patternOption= "-put|-call";
  var titleBlockStart= "<title>";
  var titleBlockEnd= "</title>";
  var pattern= null;
  
  // Regular stock quote?
  pattern= new RegExp(titleBlockStart + "\.*(" + patternStock + ")\.*" + titleBlockEnd, "i");
  if (pattern.test(data))
  {
    return ["qwidget_", "Red"];
  }
  
  // Option price quote?
  pattern= new RegExp(titleBlockStart + "\.*(" + patternOption + ")\.*" + titleBlockEnd, "i");
  if (pattern.test(data))
  {
    return ["opdetails-"];
  }
  
  // Mutual fund quote?
  pattern= new RegExp(titleBlockStart + "\.*(" + patternFund + ")\.*" + titleBlockEnd, "i");
  if (pattern.test(data))
  {
    return ["<b>"];
  }
  
  return false;
};


/**
 * ConstructUrlOptionNasdaq()
 *
 * Construct a NASDAQ quote query URL for the specified option contract
 *
 */
function ConstructUrlOptionNasdaq(option, verbose)
{
  // Declare constants and local variables
  var urlHead= "https://old.nasdaq.com/symbol/";
  var urlMid= "/option-chain/";
  var symbol= "";
  var optionSpecification= "";
  
  // Example: "https://old.nasdaq.com/symbol/aapl/option-chain/180420C00190000-aapl-call"
  
  // Extract symbol and the rest as two chunks from the option specification
  var pattern= /([a-z]+)(\d+[cp]\d+)/i;
  var chunks= pattern.exec(option);
      
  if (chunks)
  {
    // Looks like we matched and extracted the proper bits
    symbol= chunks[1].toLowerCase();
    optionSpecification= chunks[2].toLowerCase() + "-" + symbol + "-";
    pattern= /\d+c\d+/i;
    if (pattern.test(option))
    {
      // Call
      optionSpecification= optionSpecification + "call";
    }
    else
    {
      // Put
      optionSpecification= optionSpecification + "put";
    }
    
    return urlHead + symbol + urlMid + optionSpecification;
  }
  
  return null;
};


/**
 * ConstructUrlQuoteNasdaq()
 *
 * Construct a NASDAQ quote query URL for the specified symbol
 *
 */
function ConstructUrlQuoteNasdaq(symbol, verbose)
{
  // Declare constants and local variables
  var urlHead= "https://old.nasdaq.com/symbol/";
  
  // Example: "https://old.nasdaq.com/symbol/aapl"
  
  if (typeof symbol == "string")
  {
    // Convert the specific symbol to lower case and append it to the general query URL
    return urlHead + symbol.toLowerCase();
  }
  
  return null;
};