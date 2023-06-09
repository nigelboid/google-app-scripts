/**
 * RunQuotes()
 * Main entry point for the script
 *
 * Obtain and save prices for a list of symbols
 *
 */
function RunQuotes(afterHours, test)
{
  // Declare constants and local variables
  var sheetID= GetMainSheetID();
  var verbose= false;
  var success= false;
  var confirmNumbers= true;
  var limit= 0;

  var staleName= "PortfolioHeldEquitiesStaleCount";
  var stale= GetValueByName(sheetID, staleName, verbose, confirmNumbers, limit);

  if (stale > 0)
  {
    // Stale ticker symbols detected -- make a fresh copy
    if (SaveValues(sheetID, "PortfolioHeldEquitiesUniqueLookup", "PortfolioHeldEquitiesUniqueLookupSaved", verbose))
    {
      // Values updated -- update time stamp
      UpdateTime(sheetID, "PortfolioHeldEquitiesUniqueLookupSavedUpdateTime", verbose)
    }
  }
  
  if (test)
  {
    Logger.log("[RunQuotes] Testing...");
    success= RunEquitiesTest(sheetID, verbose);
  }
  else if (afterHours)
  {
    // Only check those which are sensitive after market hours
    success= RunOptionsAfterHours(sheetID, verbose);
    success= success && RunEquitiesAfterHours(sheetID, verbose);
  }
  else
  {
    // Check all standard prices
    success= RunOptions(sheetID, verbose);
    if (!success)
    {
      // Split into separate runs
      success= RunEquities(sheetID, verbose);
    }
  }
  
  LogSend(sheetID);
  return success;
};


/**
 * RunEquities()
 *
 * Obtain and save equity prices
 *
 */
function RunEquities(sheetID, verbose)
{
  // Declare constants and local variables
  var symbolsTableName= "QuotesList";
  var timeStampName= "QuotesTimeStamp";
  var checkStatusName= "QuotesCheckStatus";
  var pricesTableName= "QuotesPrices";
  var labelsTableName= "QuotesLabels";
  var updateStatusName= "QuotesUpdateStatus";
  var updateTimeName= "QuotesUpdateTime";
  var forceRefreshName= "QuotesForceRefresh";
  var pollMinutesName= "QuotesPollMinutes";
  var urlName= "QuotesURL";
  var optionPrices= false;

  return RefreshPrices(sheetID, symbolsTableName, timeStampName, checkStatusName, pricesTableName, labelsTableName,
                        updateStatusName, updateTimeName, forceRefreshName, pollMinutesName, urlName, optionPrices, verbose);
};


/**
 * RunEquitiesAfterHours()
 *
 * Obtain and save equity prices for after hours changes
 *
 */
function RunEquitiesAfterHours(sheetID, verbose)
{
  // Declare constants and local variables
  var symbolsTableName= "QuotesList";
  var timeStampName= "QuotesTimeStamp";
  var checkStatusName= "QuotesCheckStatus";
  var pricesTableName= "QuotesPrices";
  var labelsTableName= "QuotesLabels";
  var updateStatusName= "QuotesUpdateStatus";
  var updateTimeName= "QuotesUpdateTime";
  var forceRefreshName= "QuotesForceRefreshAfterHours";
  var pollMinutesName= "QuotesPollMinutes";
  var urlName= "QuotesURL";
  var optionPrices= false;
  
  if (!IsMarketOpen(sheetID, optionPrices, verbose))
  {
    return RefreshPrices(sheetID, symbolsTableName, timeStampName, checkStatusName, pricesTableName, labelsTableName,
                          updateStatusName, updateTimeName, forceRefreshName, pollMinutesName, urlName, optionPrices, verbose);
  }
  else
  {
    return false;
  }
};


/**
 * RunEquitiesTest()
 *
 * Obtain and save equity prices (for testing only)
 *
 */
function RunEquitiesTest(sheetID, verbose)
{
  // Declare constants and local variables
  var symbolsTableName= "QuotesList";
  var timeStampName= "QuotesTimeStamp";
  var checkStatusName= "QuotesCheckStatus";
  var pricesTableName= "QuotesPrices";
  var labelsTableName= "QuotesLabels";
  var updateStatusName= "QuotesUpdateStatus";
  var updateTimeName= "QuotesUpdateTime";
  var forceRefreshName= "QuotesForceRefresh";
  var pollMinutesName= "QuotesPollMinutes";
  var urlName= "QuotesURL";
  var optionPrices= false;
  var test= true;
  
  return RefreshPrices(sheetID, symbolsTableName, timeStampName, checkStatusName, pricesTableName, labelsTableName,
                        updateStatusName, updateTimeName, forceRefreshName, pollMinutesName, urlName, optionPrices, verbose, test);
};


/**
 * RunOptions()
 *
 * Obtain and save option prices
 *
 */
function RunOptions(sheetID, verbose)
{
  // Declare constants and local variables
  var symbolsTableName= "OptionsList";
  var timeStampName= "OptionsTimeStamp";
  var checkStatusName= "OptionsCheckStatus";
  var pricesTableName= "OptionsPrices";
  var labelsTableName= "OptionsLabels";
  var updateStatusName= "OptionsUpdateStatus";
  var updateTimeName= "OptionsUpdateTime";
  var forceRefreshName= "OptionsForceRefresh";
  var pollMinutesName= "OptionsPollMinutes";
  var urlName= "OptionsURL";
  var optionPrices= true;

  return RefreshPrices(sheetID, symbolsTableName, timeStampName, checkStatusName, pricesTableName, labelsTableName,
                        updateStatusName, updateTimeName, forceRefreshName, pollMinutesName, urlName, optionPrices, verbose);
};


/**
 * RunOptionsAfterHours()
 *
 * Obtain and save option prices
 *
 */
function RunOptionsAfterHours(sheetID, verbose)
{
  // Declare constants and local variables
  var symbolsTableName= "OptionsList";
  var timeStampName= "OptionsTimeStamp";
  var checkStatusName= "OptionsCheckStatus";
  var pricesTableName= "OptionsPrices";
  var labelsTableName= "OptionsLabels";
  var updateStatusName= "OptionsUpdateStatus";
  var updateTimeName= "OptionsUpdateTime";
  var forceRefreshName= "OptionsForceRefreshAfterHours";
  var pollMinutesName= "OptionsPollMinutes";
  var urlName= "OptionsURL";
  var optionPrices= true;
  
  if (!IsMarketOpen(sheetID, optionPrices, verbose))
  {
    return RefreshPrices(sheetID, symbolsTableName, timeStampName, checkStatusName, pricesTableName, labelsTableName,
                          updateStatusName, updateTimeName, forceRefreshName, pollMinutesName, urlName, optionPrices, verbose);
  }
  else
  {
    return false;
  }
};


/**
 * RunOptionsTest()
 *
 * Obtain and save option prices (for testing only)
 *
 */
function RunOptionsTest(sheetID, verbose)
{
  // Declare constants and local variables
  var symbolsTableName= "OptionsList";
  var timeStampName= "OptionsTimeStamp";
  var checkStatusName= "OptionsCheckStatus";
  var pricesTableName= "OptionsPrices";
  var labelsTableName= "OptionsLabels";
  var updateStatusName= "OptionsUpdateStatus";
  var updateTimeName= "OptionsUpdateTime";
  var forceRefreshName= "OptionsForceRefresh";
  var pollMinutesName= "OptionsPollMinutes";
  var urlName= "OptionsURL";
  var optionPrices= true;
  var test= true;
  
  Logger.log("[RunOptionsTest] Testing...");
  
  return RefreshPrices(sheetID, symbolsTableName, timeStampName, checkStatusName, pricesTableName, labelsTableName,
                        updateStatusName, updateTimeName, forceRefreshName, pollMinutesName, urlName, optionPrices, verbose, test);
};


/**
 * RefreshPrices()
 *
 * Obtain and save prices
 *
 */
function RefreshPrices(sheetID, symbolsTableName, timeStampName, checkStatusName, pricesTableName, labelsTableName,
                        updateStatusName, updateTimeName, forceRefreshName, pollMinutesName, urlName, optionPrices, verbose, test)
{
  // Declare constants and local variables
  var symbols= null;
  var labels= null;
  var priceTable= null;
  var scriptTime= new Date();
  var pollMinutes= GetValueByName(sheetID, pollMinutesName, verbose);
  var pollInterval= null;
  var pollIntervalDefault= 9;
  var minuteToMillisecondConversionFactor= 60 * 1000;
  var forceRefresh= false;
  var success= false;

  if (test)
  {
    Logger.log("[RefreshPrices] Forcing a run for testing...");
    forceRefresh= test;
  }
  else if (forceRefreshName)
  {
    // Spreadsheet refresh flag location supplied -- use spreadsheet value
    forceRefresh= GetValueByName(sheetID, forceRefreshName, verbose);
  }
 
  if (pollMinutes == undefined || pollMinutes == null)
  {
    // Set to default polling interval
    pollInterval= pollIntervalDefault * minuteToMillisecondConversionFactor;
  }
  else
  {
    // Convert to custom polling interval, as specified
    pollInterval= pollMinutes * minuteToMillisecondConversionFactor;
  }
  
  if (forceRefresh || IsMarketOpen(sheetID, optionPrices, verbose))
  {
    // Invoked during trading hours, proceed
    var timeStamp= GetValueByName(sheetID, timeStampName, verbose);
    
    if (forceRefresh || ((timeStamp + pollInterval) < scriptTime.getTime()))
    {
      // Enough time has elapsed since we last got prices, obtain current data from our option prices table
      SetValueByName(sheetID, checkStatusName, "Running... [" + DateToLocaleString(scriptTime) + "]", verbose);
      
      symbols= GetTableByNameSimple(sheetID, symbolsTableName, verbose);
      
      if (symbols)
      {
        // Extract labels from the top row
        labels= GetTableByNameSimple(sheetID, labelsTableName, verbose);
        labels= labels[0];
        
        if (labels)
        {
          // Symbol and label lists obtained, update time stamp and proceed
          if (SetValueByName(sheetID, timeStampName, scriptTime.getTime(), verbose))
          {
            // Time stamp updated to current run, get latest prices
            var urlHead= GetValueByName(sheetID, urlName, verbose);
            var prices= GetQuotes(sheetID, symbols, labels, urlHead, optionPrices, verbose, test);
            
            if (test)
            {
              Logger.log("[RefreshPrices] Test results for prices:");
              Logger.log("[RefreshPrices] %s", DumpObjectAsPrettyText(prices, "  "));
            }
            else if (prices)
            {
              if (typeof prices == "string")
              {
                // Looks like we have an error message
                Logger.log("[RefreshPrices] Could not get quotes! (%s)", prices);
                
                var updateTime= new Date();
                SetValueByName(sheetID, updateStatusName, "Could not get quotes [" + DateToLocaleString(updateTime) + "]", verbose);
                SetValueByName(sheetID, checkStatusName, "Failed [" + DateToLocaleString(scriptTime) + "]", verbose);
              }
              else
              {
                // Looks like we have prices, check time stamp
                timeStamp= GetValueByName(sheetID, timeStampName, verbose);
                  
                if (timeStamp == scriptTime.getTime())
                {
                  // Looks like nothing has clobbered our run, proceed
                  if (priceTable= GetTableByNameSimple(sheetID, pricesTableName, verbose))
                  {
                    // Previous price table obtained, apply updates
                    if (priceTable= UpdatePrices(symbols, priceTable, prices, verbose))
                    {
                      // Commit updates
                      if (SetTableByName(sheetID, pricesTableName, priceTable, verbose))
                      {
                        // Updates applied, update stamps
                        var updateTime= new Date();
                        SetValueByName(sheetID, timeStampName, updateTime.getTime(), verbose);
                        SetValueByName(sheetID, updateTimeName, DateToLocaleString(updateTime), verbose);
                        SetValueByName(sheetID, updateStatusName, "Updated [" + DateToLocaleString(updateTime) + "]", verbose);
                        SetValueByName(sheetID, checkStatusName, "Completed [" + DateToLocaleString(scriptTime) + "]", verbose);
                        
                        success= true;
                      }
                      else
                      {
                        Logger.log("[RefreshPrices] Could not write data to range named <%s> in spreadsheet ID <%s>.",
                                   pricesTableName, sheetID);
                        
                        var updateTime= new Date();
                        SetValueByName(sheetID, updateStatusName, "Could not write prices [" + DateToLocaleString(updateTime) + "]",
                                        verbose);
                        SetValueByName(sheetID, checkStatusName, "Failed [" + DateToLocaleString(scriptTime) + "]", verbose);
                      }
                    }
                    else
                    {
                      Logger.log("[RefreshPrices] Could not update transform data into form ready for writing.");
                        
                      var updateTime= new Date();
                      SetValueByName(sheetID, updateStatusName, "Could not transform prices [" + DateToLocaleString(updateTime) + "]",
                                      verbose);
                      SetValueByName(sheetID, checkStatusName, "Failed [" + DateToLocaleString(scriptTime) + "]", verbose);
                    }
                  }
                  else
                  {
                    Logger.log("[RefreshPrices] Could not read data from range named <%s> in spreadsheet ID <%s>.",
                                pricesTableName, sheetID);
                        
                    var updateTime= new Date();
                    SetValueByName(sheetID, updateStatusName, "Could not read current snapshot [" + DateToLocaleString(updateTime) + "]",
                                    verbose);
                    SetValueByName(sheetID, checkStatusName, "Failed [" + DateToLocaleString(scriptTime) + "]", verbose);
                  }
                }
                else
                {
                  Logger.log("[RefreshPrices] Superseded by another run (%s minutes later).",
                             ((timeStamp - scriptTime.getTime()) / minuteToMillisecondConversionFactor).toFixed(2));
                }
              }
            }
            else
            {
              Logger.log("[RefreshPrices] Could not get price data.");
              
              var updateTime= new Date();
              SetValueByName(sheetID, updateStatusName, "Could not get prices [" + DateToLocaleString(updateTime) + "]", verbose);
              SetValueByName(sheetID, checkStatusName, "Failed [" + DateToLocaleString(scriptTime) + "]", verbose);
            }
          }
          else
          {
            Logger.log("[RefreshPrices] Could not update time stamp <%s> in spreadsheet ID <%s>.", timeStampName, sheetID);
              
            var updateTime= new Date();
            SetValueByName(sheetID, updateStatusName, "Could not update time stamp [" + DateToLocaleString(updateTime) + "]", verbose);
            SetValueByName(sheetID, checkStatusName, "Failed [" + DateToLocaleString(scriptTime) + "]", verbose);
          }
        }
        else
        {
          Logger.log("[RefreshPrices] Could not read data from range named <%s> in spreadsheet ID <%s>.", labelsTableName, sheetID);
              
          var updateTime= new Date();
          SetValueByName(sheetID, updateStatusName, "Could not obtain list of labels [" + DateToLocaleString(updateTime) + "]", verbose);
          SetValueByName(sheetID, checkStatusName, "Failed [" + DateToLocaleString(scriptTime) + "]", verbose);
        }
      }
      else
      {
        Logger.log("[RefreshPrices] Could not read data from range named <%s> in spreadsheet ID <%s>.", symbolsTableName, sheetID);
            
        var updateTime= new Date();
        SetValueByName(sheetID, updateStatusName, "Could not obtain list of symbols [" + DateToLocaleString(updateTime) + "]", verbose);
        SetValueByName(sheetID, checkStatusName, "Failed [" + DateToLocaleString(scriptTime) + "]", verbose);
      }
    }
    else
    {
      // Update status
      var minutes= (pollInterval-(scriptTime.getTime()-timeStamp)) / minuteToMillisecondConversionFactor;
      SetValueByName(sheetID, checkStatusName, "Invoked too soon (" + minutes.toFixed(2)
      + " minutes remaining) [" + DateToLocaleString(scriptTime) + "]", verbose);
    }
  }
  else
  {
    // Update status
    SetValueByName(sheetID, checkStatusName, "Market Closed [" + DateToLocaleString(scriptTime) + "]", verbose);
  }
  
  return success;
};


/**
 * GetQuotes()
 *
 * Obtain prices from a specific service with a specific method
 *
 */
function GetQuotes(id, symbols, labels, urlHead, optionPrices, verbose, test)
{

  if (test)
  {
    Logger.log("[GetQuotes] Still testing...");
    return GetQuotesTDA(id, symbols, labels, urlHead, verbose);
  }
  else if (optionPrices)
  {
    return GetQuotesTDA(id, symbols, labels, urlHead, verbose);
    //return GetQuotesTradierAll(symbols, labels, urlHead, optionPrices, verbose);
  }
  else
  {
    return GetQuotesTDA(id, symbols, labels, urlHead, verbose);
    //return GetQuotesCNBCAll(symbols, labels, urlHead, optionPrices, verbose);
  }
};


/**
 * UpdatePrices()
 *
 * Save updated prices
 *
 */
function UpdatePrices(symbols, pricesTable, prices, verbose)
{
  // Declare constants and local variables
  var headerRow= 0;
  var firstDataRow= 1;
  var symbolColumn= 0;
  var labelURL= "URL";
  var labelDebug= "Debug";
  var value= null;
  var label= "";
  
  for (var vIndex= firstDataRow; vIndex < pricesTable.length; vIndex++)
  {
    if (symbols[vIndex][symbolColumn].length > 0)
    {
      // Consider each symbol
      
      for (const hIndex in pricesTable[vIndex])
      {
        // Update each entry for the given symbol
        label= pricesTable[headerRow][hIndex];
    
        if (label == labelURL || label == labelDebug)
        {
           if (prices[symbols[vIndex][symbolColumn]] == undefined)
           {
             // No entry for likely invalid symbol
             //pricesTable[vIndex][hIndex]= "";
             pricesTable[vIndex][hIndex]= symbols[vIndex][symbolColumn];
           }
          else
          {
            // Leave links as they are
            pricesTable[vIndex][hIndex]= prices[symbols[vIndex][symbolColumn]][label];
          }
        }
        else
        {
          // Transform price entries
          if (prices[symbols[vIndex][symbolColumn]] == undefined)
          {
            // No entry for likely invalid symbol
            pricesTable[vIndex][hIndex]= "no data";
          }
          else
          {
            value= prices[symbols[vIndex][symbolColumn]][label];
            if (isNaN(value))
            {
              // Garbage?! Transform into a valid link to the quote page
              pricesTable[vIndex][hIndex]= "=hyperlink(\"" + prices[symbols[vIndex][symbolColumn]][labelURL] + "\", \"" + value + "\")";
            }
            else
            {
              // Transform valid price into link to the quote page
              pricesTable[vIndex][hIndex]= "=hyperlink(\"" + prices[symbols[vIndex][symbolColumn]][labelURL] + "\", " + value + ")";
            }
          }
        }
      }
    }
    else
    {
      // No entry for this row -- clear it
      for (const hIndex in pricesTable[vIndex])
      {
        pricesTable[vIndex][hIndex]= "";
      }
    }
  }
          
  // Return updated table
  return pricesTable;
};


/**
 * IsMarketOpen()
 *
 * Is this script running during market hours?
 *
 */
function IsMarketOpen(id, optionPrices, verbose)
{
  // Declare constants and local variables
  var time= new Date();
  var marketSunday= 0;
  var marketSaturday= 6;
  var marketTimeOpen= null;
  var marketTimeClose= null;
  
  
  if (optionPrices)
  {
    // No  extended hours trading
    marketTimeOpen= GetValueByName(id, "ParameterMarketTimeOptionsOpen", verbose);
    marketTimeClose= GetValueByName(id, "ParameterMarketTimeOptionsClose", verbose);
  }
  else
  {
    // Extended hours trading
    marketTimeOpen= GetValueByName(id, "ParameterMarketTimeOpen", verbose);
    marketTimeClose= GetValueByName(id, "ParameterMarketTimeClose", verbose);
  }
  
  if (marketTimeOpen == undefined || marketTimeOpen == null)
  {
    // No parameter provided? Use default instead
    marketTimeOpen= 600;
  }
  
  if (marketTimeClose == undefined || marketTimeClose == null)
  {
    // No parameter provided? Use default instead
    marketTimeClose= 1900;
  }
  
  return (time.getDay() > marketSunday && time.getDay() < marketSaturday) &&
    ((time.getHours() * 100 + time.getMinutes()) > marketTimeOpen && (time.getHours() * 100 + time.getMinutes()) < marketTimeClose);
};


/**
 * ConstructUrlQuote()
 *
 * Construct a quote query URL for the specified symbol (default to Yahoo! Finance)
 *
 */
function ConstructUrlQuote(symbol, urlHead, verbose)
{
  if (!urlHead)
  {
    // Nothing supplied -- use Yahoo! Finance as default
    urlHead= "https://finance.yahoo.com/quote/";
    
    // Example: "https://finance.yahoo.com/quote/aapl"
  }
  
  if (typeof symbol == "string")
  {
    // Convert the specific symbol to lower case and append it to the general query URL
    symbol= symbol.toLowerCase();
    
    return urlHead + symbol;
  }
  
  return null;
};