/**
 * RunBoxTradeCandidates()
 *
 * Find suitable contract expirations for box trades
 *
 */
function RunBoxTradeCandidates(afterHours)
{
  // Declare constants and local variables
  const sheetID = GetMainSheetID();
  const optionPrices = true;
  const verbose = false;
  var success = false;
  
  const nextUpdateTime = GetValueByName(sheetID, "IndexStranglesBoxesUpdateNext", verbose);
  const currentTime = new Date();

  if (afterHours == undefined)
  {
    afterHours = true;
  }
  
  if (IsMarketOpen(sheetID, optionPrices, verbose) && (currentTime > nextUpdateTime) || afterHours)
  {
    // Looks like we are due for an update
    const parameters = GetBoxParameters(sheetID, verbose);
    const candidates = FindBestBoxTradeCandidates(parameters);

    if (candidates.length > 0)
    {
      // Looks like we obtained candidate boxes -- commit them to a table
      if (SetTableByName(sheetID, "IndexStranglesBoxes", candidates, verbose))
      {
        UpdateTime(sheetID, "IndexStranglesBoxesUpdateTime", verbose);
        SetValueByName(sheetID, "IndexStranglesBoxesUpdateStatus", "Updated [" + DateToLocaleString(currentTime) + "]", verbose);

        success = true;
      }
      else
      {
        SetValueByName(sheetID, "IndexStranglesBoxesUpdateStatus",
                        "Failed to set new candidates [" + DateToLocaleString(currentTime) + "]", verbose);

        Log(`Failed to update table named 'IndexStranglesCandidates' in sheet <${sheetID}> with candidates: <${candidates}>!`);
      }
    }
    else
    {
      // No candidates found?
      SetValueByName(sheetID, "IndexStranglesBoxesUpdateStatus",
                      "Failed to find new candidates [" + DateToLocaleString(currentTime) + "]", verbose);

      Log(`Found no candidates <${candidates}>!`);
      success = true;
    }
  }
  
  LogSend(sheetID);
  return success;
};


/**
 * FindBestBoxTradeCandidates()
 *
 * Find highest yielding box trades
 *
 */
function FindBestBoxTradeCandidates(parameters)
{
  // Declare constants and local variables
  const labelPuts = "putExpDateMap";
  const labelCalls = "callExpDateMap";
  
  var candidates = [];
  var top = parameters["top"];

  const response = GetChainForSymbolByExpirationSchwab(parameters["id"], parameters["underlying"], parameters["dte_earliest"],
                                                        parameters["dte_latest"], parameters["verbose"]);

  if (typeof response == "string")
  {
    // Looks like we have an error message
    Log(`Could not get contracts! (${response})`);
  }
  else if (response)
  {
    // Data fetched -- extract
    const contentParsed = ExtractContentSchwab(response);
    const locale = "en-US";
    const options =
    {
      style : "percent",
      maximumFractionDigits : 7,
      minimumFractionDigits : 2
    }

    if (contentParsed)
    {
      // Find highest yielding box trade for a given invested amount
      var ytm = null;
      var yields = FindBoxTradeYields(contentParsed[labelPuts], contentParsed[labelCalls],
                                      parameters["amount"], parameters["commission"], parameters["iterations"]);

      yieldsOrdered = Object.keys(yields).sort();

      while (top)
      {
        // Preserve top results
        ytm = yieldsOrdered.pop();

        candidates.push([yields[ytm]["putHigh"]["contract"],
                          yields[ytm]["putHigh"]["price"], yields[ytm]["amount"], parseFloat(ytm).toLocaleString(locale, options)]);
        candidates.push([yields[ytm]["callHigh"]["contract"],
                          yields[ytm]["callHigh"]["price"], yields[ytm]["amount"], parseFloat(ytm).toLocaleString(locale, options)]);
        candidates.push([yields[ytm]["putLow"]["contract"],
                          yields[ytm]["putLow"]["price"], yields[ytm]["amount"], parseFloat(ytm).toLocaleString(locale, options)]);
        candidates.push([yields[ytm]["callLow"]["contract"],
                          yields[ytm]["callLow"]["price"], yields[ytm]["amount"], parseFloat(ytm).toLocaleString(locale, options)]);

        top--;
      }
    }
    else
    {
      // Did we get bad content?
      Log(`Could not parse returned content: <${contentParsed}>`);
    }
  }
  else
  {
    // Failed to fetch results
    Log(`Could not fetch option chain for <${underlying}> for expirations`
        + ` between <${dteEarliest.toFixed(0)}> and <${dteLatest.toFixed(0)}> days!`);
  }

  return candidates;
};


/**
 * FindBoxTradeYields()
 *
 * Return a list of box trade contracts with a viable yield for a given invested amount
 *
 */
function FindBoxTradeYields(puts, calls, amountMaximum, commission, iterationsMaximum)
{
  // Declare constants and local variables
  const expirations = Object.keys(puts).sort();
  const expirationDelimiter = ":";
  const preferWeekly = false;
  var yields = {};

  for (const expiration of expirations)
  {
    // Collect yields from contracts which statisfy our box value (amount) for each expiration (DTE)
    const dte = expiration.split(expirationDelimiter)[1];
    const contractsPuts = puts[expiration];
    const contractsCalls = calls[expiration];
    const strikes = Object.keys(contractsPuts).sort();
    var strikeHigh = null;
    var iterations = iterationsMaximum;
    var amount = amountMaximum;

    while (iterations)
    {
      // Search for successively halved amounts for a number of specified iterations
      if (contractsCalls != undefined)
      {
        // Calls also exist for this expiration
        
        for (const strikeLow of strikes)
        {
          // Collect yields from contracts whcih satisfy our box value (amount) for this expiration (DTE)
          strikeHigh= (parseFloat(strikeLow) + amount).toFixed(1).toString();

          if (contractsPuts[strikeHigh] != undefined && contractsCalls[strikeLow] != undefined && contractsCalls[strikeHigh] != undefined)
          {
            // Our box ostensibly exists -- find and confirm weekly contracts
            const pricePutLow = ExtractPriceSchwab(contractsPuts[strikeLow], amount, preferWeekly);
            const pricePutHigh = ExtractPriceSchwab(contractsPuts[strikeHigh], amount, preferWeekly);
            const priceCallLow = ExtractPriceSchwab(contractsCalls[strikeLow], amount, preferWeekly);
            const priceCallHigh = ExtractPriceSchwab(contractsCalls[strikeHigh], amount, preferWeekly);
            const yieldToMaturity = BoxYield(pricePutLow, pricePutHigh, priceCallLow, priceCallHigh, dte, amount, commission);

            if (yieldToMaturity)
            {
              // We have a viable yield, add it to our list
              
              yields[yieldToMaturity] =
              {
                putLow : pricePutLow,
                putHigh : pricePutHigh,
                callLow : priceCallLow,
                callHigh : priceCallHigh,
                amount : amount
              };
            }
          }
        }
      }

      // Prepare for next iteration
      iterations--;
      amount = amount / 2;
    }
  }

  return yields;
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
  const labelSymbol = "symbol";
  const labelLastPrice = "last";
  const labelClosePrice = "closePrice";
  const labelBidPrice = "bid";
  const labelAskPrice = "ask";
  const labelDelta = "delta";
  const deltaBad = "-999.0";
  const symbolDelimiter = " ";
  const weekly = "W";
  var price = null;

  if (preferWeekly == undefined)
  {
    // Usually, prefer weekly expirations
    preferWeekly = true;
  }

  for (quote of quotes)
  {
    if (quote[labelDelta] != deltaBad)
    {
      const last = parseFloat(quote[labelLastPrice]);
      const close = parseFloat(quote[labelClosePrice]);
      const bid = parseFloat(quote[labelBidPrice]);
      const ask = parseFloat(quote[labelAskPrice]);

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
          price = {price : close, contract : quote[labelSymbol]};
        }
        else
        {
          // Use the mid point of the spread for price
          price = {price : bid + (ask - bid) / 2, contract : quote[labelSymbol]};
        }
      }

      if (price > amount)
      {
        // Avoid extreme strikes
        price = null;
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


/**
 * BoxYield()
 *
 * Compute yield for a box spread
 *
 */
function BoxYield(pricePutLow, pricePutHigh, priceCallLow, priceCallHigh, dte, amount, commission)
{
  // Declare constants and local variables
  const daysPerYear = 360;
  var yieldToMaturity = null;

  commission = commission * 4 / 100;

  if (pricePutLow != undefined && pricePutHigh != undefined && priceCallLow != undefined && priceCallHigh != undefined)
  {
    // We appear to have real prices -- compute yield
    const cost = pricePutHigh["price"] - priceCallHigh["price"] - pricePutLow["price"] + priceCallLow["price"] + commission;

    yieldToMaturity = Math.pow(amount / cost, daysPerYear / dte) - 1;

    if (isNaN(yieldToMaturity) || yieldToMaturity < 0)
    {
      yieldToMaturity = null;
    }
  }

  return yieldToMaturity;
};


/**
 * GetBoxParameters()
 *
 * Load specified parameters and assign default values to missing parameters
 *
 */
function GetBoxParameters(sheetID, verbose)
{
  // Declare constants and local variables
  const underlyingDefault = "$SPX";
  const dteEarliestDefault = 90;
  const dteLatestDefault = 360;
  const amountDefault = 1000;
  const topDefault = 3;
  const iterationsDefault = 3;
  
  var parameters= GetParameters(sheetID, "IndexStranglesBoxesParameters", verbose);

  if (parameters["underlying"] == undefined)
  {
    parameters["underlying"] = underlyingDefault;
  }

  if (parameters["dte_earliest"] == undefined)
  {
    parameters["dte_earliest"] = dteEarliestDefault;
  }

  if (parameters["dte_latest"] == undefined)
  {
    parameters["dte_latest"] = dteLatestDefault;
  }

  if (parameters["amount"] == undefined)
  {
    parameters["amount"] = amountDefault;
  }

  if (parameters["iterations"] == undefined)
  {
    parameters["iterations"] = iterationsDefault;
  }

  if (parameters["top"] == undefined)
  {
    parameters["top"] = topDefault;
  }

  return parameters;
};