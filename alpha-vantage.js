/**
 * RunQuotesAlphaVantage()
 *
 * Obtain and save prices for a list of symbols from Alpha Vantage
 *
 */
function RunQuotesAlphaVantage(sheetID, verbose)
{
  // Declare constants and local variables
  var urlName= "AlphaVantageURL";
  var symbolListTableName= "AlphaVantageSymbolList";
  var timeStampName= "AlphaVantageTimeStamp";
  var checkStatusName= "AlphaVantageCheckStatus";
  var quotesTableName= "AlphaVantageQuotes";
  var updateStatusName= "AlphaVantageUpdateStatus";
  var forceRefreshName= "AlphaVantageForceRefresh";
  var scriptTime= new Date();
  var minuteToMillisecondConversion= 60 * 1000;
  var pollIntervalMinutes= 4;
  var pollInterval= pollIntervalMinutes * minuteToMillisecondConversion;
  var marketSunday= 0;
  var marketSaturday= 6;
  var marketTimeOpen= 830;
  var marketTimeClose= 1559;
  
  var forceRefresh= GetValueByName(sheetID, forceRefreshName, verbose);
  
  //pollInterval= 0 * minuteToMillisecondConversion;
  //marketSunday= -1;
  //marketSaturday= 7;
  //marketTimeOpen= 0;
  //marketTimeClose= 2400;
  
  //Logger.log("[RunQuotesAlphaVantage] Starting...");
  //Logger.log(prices);
  //LogSend(sheetID);
  
  if (forceRefresh || ((scriptTime.getDay() > marketSunday && scriptTime.getDay() < marketSaturday)
  && ((scriptTime.getHours() * 100 + scriptTime.getMinutes()) > marketTimeOpen
  && (scriptTime.getHours() * 100 + scriptTime.getMinutes()) < marketTimeClose)))
  {
    // Invoked during trading hours (or forcing an update), proceed
    var timeStamp= GetValueByName(sheetID, timeStampName, verbose);
    
    if (forceRefresh || ((timeStamp + pollInterval) < scriptTime.getTime()))
    {
      // Enough time has elapsed since we last got prices (or forcing an update), obtain current data from our option prices table
      
      SetValueByName(sheetID, checkStatusName, "Running... [" + scriptTime.toLocaleString() + "]", verbose);
      var symbols= GetTableByName(sheetID, symbolListTableName, 0, false, 0, false, verbose);
      if (symbols)
      {
        // Options list obtained, update time stamp and proceed
        if (SetValueByName(sheetID, timeStampName, scriptTime.getTime(), verbose))
        {
          // Time stamp updated to current run, get URL for quotes
          var url= GetValueByName(sheetID, urlName, verbose);
          
          if (url)
          {
            //Logger.log("[RunQuotesAlphaVantage] URL:");
            //Logger.log(url);
            //LogSend(sheetID);
              
            // URL obtained, get latest prices
            var prices= ExtractDataAlphaVantageBatch(UrlFetchApp.fetch(url), verbose);
            
            if (prices)
            {
              // Looks like we have prices, check time stamp
              timeStamp= GetValueByName(sheetID, timeStampName, verbose);
                
              if (timeStamp == scriptTime.getTime())
              {
                // Looks like nothing has clobbered our run, proceed
                var pricesTable= GetTableByName(sheetID, quotesTableName, 0, false, 0, false, verbose);
                if (pricesTable)
                {
                  // Previous price table obtained, apply updates
                  if (pricesTable= UpdatePricesAlphaVantage(symbols, pricesTable, prices, verbose))
                  {
                    // Commit updates
                    if (SetTableByName(sheetID, quotesTableName, pricesTable, verbose))
                    {
                      // Updates applied, update stamps
                      var updateTime= new Date();
                      SetValueByName(sheetID, timeStampName, updateTime.getTime(), verbose);
                      SetValueByName(sheetID, updateStatusName, "Updated [" + updateTime.toLocaleString() + "]", verbose);
                      SetValueByName(sheetID, checkStatusName, "Completed [" + scriptTime.toLocaleString() + "]", verbose);
                    }
                    else
                    {
                      Logger.log("[RunQuotesAlphaVantage] Could not write data to range named <%s> in spreadsheet ID <%s>.", quotesTableName, sheetID);
                      
                      var updateTime= new Date();
                      SetValueByName(sheetID, updateStatusName, "Could not write prices [" + updateTime.toLocaleString() + "]", verbose);
                      SetValueByName(sheetID, checkStatusName, "Failed [" + scriptTime.toLocaleString() + "]", verbose);
                    }
                  }
                  else
                  {
                    Logger.log("[RunQuotesAlphaVantage] Could not transform data into form ready for writing.");
                      
                    var updateTime= new Date();
                    SetValueByName(sheetID, updateStatusName, "Could not transform prices [" + updateTime.toLocaleString() + "]", verbose);
                    SetValueByName(sheetID, checkStatusName, "Failed [" + scriptTime.toLocaleString() + "]", verbose);
                  }
                }
                else
                {
                  var updateTime= new Date();
                  SetValueByName(sheetID, updateStatusName, "Could not read current snapshot [" + updateTime.toLocaleString() + "]", verbose);
                  SetValueByName(sheetID, checkStatusName, "Failed [" + scriptTime.toLocaleString() + "]", verbose);
                }
              }
              else
              {
                Logger.log("[RunQuotesAlphaVantage] Superseded by another run (%s minutes later).",
                           ((timeStamp - scriptTime.getTime()) / minuteToMillisecondConversion).toFixed(2));
              }
            }
            else
            {
              if (verbose)
              {
                Logger.log("[RunQuotesAlphaVantage] Could not get price data.");
              }
              
              var updateTime= new Date();
              SetValueByName(sheetID, updateStatusName, "Could not get prices [" + updateTime.toLocaleString() + "]", verbose);
              SetValueByName(sheetID, checkStatusName, "Failed [" + scriptTime.toLocaleString() + "]", verbose);
            }
          }
          else
          {
            Logger.log("[RunQuotesAlphaVantage] Could not get URL.");
              
            var updateTime= new Date();
            SetValueByName(sheetID, updateStatusName, "Could not get URL [" + updateTime.toLocaleString() + "]", verbose);
            SetValueByName(sheetID, checkStatusName, "Failed [" + scriptTime.toLocaleString() + "]", verbose);
          }
        }
        else
        {
          Logger.log("[RunQuotesAlphaVantage] Could not update time stamp <%s> in spreadsheet ID <%s>.", timeStampName, sheetID);
        }
      }
      else
      {
        Logger.log("[RunQuotesAlphaVantage] Could not read data from range named <%s> in spreadsheet ID <%s>.", symbolListTableName, sheetID);
      }
    }
    else
    {
      // Update status
      var minutes= (pollInterval-(scriptTime.getTime()-timeStamp)) / minuteToMillisecondConversion;
      SetValueByName(sheetID, checkStatusName,
                     "Invoked too soon (" + minutes.toFixed(2) + " minutes remaining) [" + scriptTime.toLocaleString() + "]", verbose);
    }
  }
  else
  {
    // Update status
    SetValueByName(sheetID, checkStatusName, "Market Closed [" + scriptTime.toLocaleString() + "]", verbose);
  }
};


/**
 * ExtractDataAlphaVantageBatch()
 *
 * Extract quotes from obtained Alpha Vantage response
 *
 */
function ExtractDataAlphaVantageBatch(response, verbose)
{
  // Declare constants and local variables
  var typeSymbol= "1. symbol";
  var labelSymbol= "symbol";
  var types= ["2. price", "3. volume", "4. timestamp"];
  var labelQuotes= "Stock Quotes";
  var firstMatch= 1;
  var responseOK= 200;
  var prices= {};
  
  //Logger.log("[ExtractDataAlphaVantage] Alpha Vantage response:");
  //Logger.log(response);
  //LogSend(GetMainSheetID());
  
  if (response.getResponseCode() == responseOK)
  {
    // Looks like we have a valid data response
    var content= response.getContentText();
    var quotes= JSON.parse(content)[labelQuotes];
    
    if (quotes)
    {
      // Looks like we have valid quotes
      for (var quote in quotes)
      {
        // Process each returned quote
        if (quotes[quote][typeSymbol] != undefined)
        {
          // Symbol exists
          prices[quotes[quote][typeSymbol]]= {};
          prices[quotes[quote][typeSymbol]][labelSymbol]= quotes[quote][typeSymbol];
          for (var type in types)
          {
            // Drop the stupid leading number from each type label
            var label= types[type].split(" ")[1];
            
            if (quotes[quote][types[type]] == undefined)
            {
              prices[quotes[quote][typeSymbol]][label]= "no data";
            }
            else
            {
              prices[quotes[quote][typeSymbol]][label]= quotes[quote][types[type]];
            }
          }
        }
      }
    }
    else
    {
      //Logger.log("[ExtractDataAlphaVantage] Data query did not return valid quotes.");
      //Logger.log(content);
      //LogSend(GetMainSheetID());
      
      return false;
    }
  }
  else
  {
    Logger.log("[ExtractDataAlphaVantage] Data query returned error code <%s>", response.getResponseCode());
    Logger.log(response.getAllHeaders().toSource());
    //LogSend(GetMainSheetID());
    
    return false;
  }
  
  return prices;
};


/**
 * UpdatePricesAlphaVantage()
 *
 * Save prices for a list of symbols
 *
 */
function UpdatePricesAlphaVantage(symbols, pricesTable, prices, verbose)
{
  // Declare constants and local variables
  var labelSymbol= "symbol";
  var headerRow= 0;
  var firstDataRow= 1;
  var symbolColumn= 0;
  var value= null;
  
  for (var vIndex= firstDataRow; vIndex < symbols.length; vIndex++)
  {
    if ((symbols[vIndex][symbolColumn].length > 0) || (prices[symbols[vIndex][symbolColumn]] != undefined))
    {
      // We have data for this row
      for (var hIndex in pricesTable[vIndex])
      {
        // Update each entry for the given symbol
        if (pricesTable[headerRow][hIndex].toLowerCase() == labelSymbol)
        {
          pricesTable[vIndex][hIndex]= prices[symbols[vIndex][symbolColumn]][labelSymbol];
        }
        else
        {
          pricesTable[vIndex][hIndex]= prices[symbols[vIndex][symbolColumn]][pricesTable[headerRow][hIndex].toLowerCase()];
        }
      }
    }
    else
    {
      // No entry for this row -- clear it
      for (var hIndex in pricesTable[vIndex])
      {
        pricesTable[vIndex][hIndex]= "";
      }
    }
  }
          
  // Return updated table
  return pricesTable;
};