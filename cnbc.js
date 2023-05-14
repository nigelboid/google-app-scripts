/**
 * GetQuotesCNBCAll()
 *
 * Obtain option prices from CNBC via parallel calls
 *
 */
function GetQuotesCNBCAll(symbols, labels, urlHead, options, verbose)
{
  // Declare constants and local variables
  var firstDataRow= 1;
  var symbolColumn= 0;
  var labelURL= "URL";
  var url= null;
  var urls= null;
  var requests= {};
  var responses= null;
  var prices= {};
  
  for (var vIndex= firstDataRow; vIndex < symbols.length; vIndex++)
  {
    // Compile a list of unique URLs
    if (symbols[vIndex][symbolColumn].length > 0)
    {
      // Live line
      url= ConstructUrlQuoteCNBC(symbols[vIndex][symbolColumn], verbose);
      
      if (url)
      {
        // Create entries for each URL in two mirroring maps for future reconciliation
        requests[url]= symbols[vIndex][symbolColumn];
      }
    }
  }
  
  // Convert unique URL keys into a plain array and fetch them in parallel
  if (urls= Object.keys(requests))
  {
    responses= FetchCNBCAll(urls, verbose);
    if (responses)
    {
      // Pages fetched -- extract data
      for (var index in responses)
      {
        // Extract data from each response
        prices[requests[urls[index]]]= ExtractDataCNBC(requests[urls[index]], responses[index], labels, verbose);
        
        // Append the URL
        prices[requests[urls[index]]][labelURL]= urls[index];
      }
    }
    else
    {
      // Failed to fetch web pages
      return "[GetQuotesCNBCAll] Could not fetch quotes!";
    }
  }
  else
  {
    // No prices to fetch?
    return "[GetQuotesCNBCAll] Could not compile queries.";
  }
  
  return prices;
};


/**
 * ExtractDataCNBC()
 *
 * Extract pricing data from CNBC quote pages (HTML source)
 *
 */
function ExtractDataCNBC(symbol, response, labels, verbose)
{
  // Declare constants and local variables
  var firstMatch= 1;
  var responseOK= 200;
  var labelDebug= "Debug";
  var values= {};
  var chunks= [];
  var patternValues= null;
  var data= null;
  var typeLabel= "type";
  var type= "";
  
  //labelDebug= undefined;
  
  if (response.getResponseCode() == responseOK)
  {
    // Looks like we have a valid data response
    data= CleanDataCNBCAccumulator(response.getContentText(), verbose);
    
    // Glean asset type
    patternValues= new RegExp('\\"' + typeLabel + '\\":\\"([A-Z]*)\\"', 'm');
    chunks= patternValues.exec(data);
    if (chunks)
    {
      // We have a match!
      type= chunks[firstMatch];
    }
    
    for (var label in labels)
    {
      // Extract each desired value from the quote
      if (label == labelDebug)
      {
        Logger.log("[ExtractDataCNBC] Why did we match <%s> in <%s>?", label, labels);
      }
      
      patternValues= new RegExp('\\"' + labels[label].toLowerCase() + '\\":\\"([-\\+]?\\d*,?\\d+\\.?\\d*)', 'im');
        
      chunks= patternValues.exec(data);
      if (chunks)
      {
        // We have a numeric match!
        //  excise all commas and the leading plus sign
        values[labels[label]]= chunks[firstMatch].replace("+", "");
      }
      else
      {
        // Failed to match a number -- try for "unch"
        patternValues= new RegExp('\\"' + labels[label].toLowerCase() + '\\":\\"unch\\"', 'im');
        if (patternValues.test(data))
        {
          // Found a zero ("unch") value
          values[labels[label]]= 0;
        }
        else
        {
          // No match?!
          values[labels[label]]= "no data";
        }
      }
    }
    
    // Capture data if the change and change % do not match
    if (labelDebug != undefined)
    {
      var last= "Last";
      var change= "Change";
      var changeP= "Change_Pct";
      var typeBond= "BOND";
      
      if (type == typeBond)
      {
        // Bond quotes return bogus change percentage
        values[labelDebug]= "Bond: ".concat("last= ", values[last], "%, ∆= ", values[change], "%");
        values[changeP]= "no data";
      }
      else if (Math.abs(values[change] / (values[last] - values[change]) * 100 - values[changeP]) > 0.1)
      {
        // All other asset types with mismatched calculated and quted values
        values[labelDebug]= "Mismatch! ".concat("last= $", values[last], ", ∆= $", values[change], ", %∆= ", values[changeP], "%, type=", type);
        
        if (verbose)
        {
          Logger.log("[ExtractDataCNBC] Data query for symbol <%s> returned mismatched data (last: $%s, ∆= $%s, %∆= %s%, computed %∆= %s%, type= %s).",
                     symbol, values[last], values[change], values[changeP],
                     (values[change] / (values[last] - values[change]) * 100).toFixed(2), type);
          Logger.log("[ExtractDataCNBC] Data:\n\n <%s>", data);
        }
      }
      else
      {
        // All other asset types with matching calculated and quted values
        values[labelDebug]= "Match: ".concat("last= $", values[last], ", ∆= $", values[change], ", %∆= ", values[changeP], "%, raw: ", data);
        //values[labelDebug]= data;
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
    
    if (verbose)
    {
      // Only report fetch errors if explicitly requested
      Logger.log("[ExtractDataCNBC] Data query for symbol <%s> returned error code <%s>.", symbol, response.getResponseCode().toFixed(0));
      //Logger.log(response.getAllHeaders());
      //Logger.log(response.getContentText());
    }
  }
  
  return values;
};


/**
 * CleanDataCNBCRegExp()
 *
 * Prepare data for extraction by reducing and cleaning the original full text with regular expressions
 *
 */
function CleanDataCNBCRegExp(data, verbose)
{
  var patternAnyBraces= new RegExp('\\{[^\\{\\}]*\\}', 'gim');
  var patternValueBraces= new RegExp('var symbolInfo = \\{([^\\{\\}]*)\\}', 'im');
  var chunks= null;
  var firstMatch= 1;
  
  // Drop conflicting information from the Extended Market and Pre-Market quotes
  //   by finding everything within curly brackets
  while (data.match(patternAnyBraces))
  {
    chunks= patternValueBraces.exec(data);
    if (chunks)
    {
      // Found clean value block
      data= chunks[firstMatch];
      break;
    }
    else
    {
      // Excise everything in braces that is not our value block
      data= data.replace(patternAnyBraces, '');
    }
  }
  
  return data;
};


/**
 * CleanDataCNBCAccumulator()
 *
 * Prepare data for extraction by reducing and cleaning the original full text by iterative interesting text accumulation
 *
 */
function CleanDataCNBCAccumulator(data, verbose)
{
  var labelValueBlock= "var symbolInfo = {";
  var labelValueBlock= '"quote":{"data":[{';
  var labelDelimiterOpen= "{";
  var labelDelimiterClose= "}";
  var character= "";
  var bottomLevel= 1;
  var cleanData= "";
  var position= data.indexOf(labelValueBlock);
  
  if (position > 0)
  {
    // Found our values label
    var level= bottomLevel;
    var startPosition= 0;
    var cleanSpan= 0;
    
    // Advance past our label
    position+= labelValueBlock.length;
    startPosition= position;
  
    while (level >= bottomLevel && position < data.length)
    {
      // Check each character while within our value block
      character= data.charAt(position);
      if (character == labelDelimiterOpen)
      {
        // Entering a higher level
        level++;
        if (cleanSpan > 0)
        {
          // Commit accumulate viable text
          cleanData+= data.substr(startPosition, cleanSpan);
          cleanSpan= 0;
        }
      }
      else if (character == labelDelimiterClose)
      {
        // Entering a lower level
        level--;
        if (level == bottomLevel)
        {
          // Start counting text after the bracket since we have reached the top level again
          startPosition= position + 1;
        }
      }
      else if (level == bottomLevel)
      {
        // Increase useful text span while at the top level
        cleanSpan++;
      }
      
      position++;
    }
    
    if (level <= bottomLevel && cleanSpan > 0)
    {
      // Append the final clean snippet
      cleanData+= data.substr(startPosition, cleanSpan);
    }
  }
  
  return cleanData;
};


/**
 * ConstructUrlQuoteCNBC()
 *
 * Construct a CNBC quote query URL for the specified symbol
 *
 */
function ConstructUrlQuoteCNBC(symbol, verbose)
{
  // Declare constants and local variables
  var urlHead= "https://www.cnbc.com/quotes/";
  
  // Example: "https://www.cnbc.com/quotes/aapl"
  
  if (typeof symbol == "string")
  {
    // Convert the specific symbol to lower case and append it to the general query URL
    return urlHead + symbol.toLowerCase();
  }
  
  return null;
};


/**
 * FetchCNBCAll()
 *
 * Fetch data from CNBC
 *
 */
function FetchCNBCAll(urls, verbose)
{
  // Declare constants and local variables
  var requests= [];
  var responses= null;
  
  // Reconstitute our list of URLs as a list of objects with URLs and parameters
  for (var index= 0; index < urls.length; index++)
  {
    requests.push({"url": urls[index], "muteHttpExceptions": true});
  }
  
  try
  {
    responses= UrlFetchApp.fetchAll(requests);
  }
  catch (error)
  {
    if (verbose)
    {
      Logger.log("[FetchCNBCAll] UrlFetchApp failed:\n".concat(error));
    }
    return null;
  }
    
  return responses;
};