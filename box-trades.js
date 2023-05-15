/**
 * RunBoxTradeCandidates()
 *
 * Find suitable contract expirations for box trades
 *
 */
function RunBoxTradeCandidates(backupRun)
{
  // Declare constants and local variables
  const sheetID= GetMainSheetID();
  const verbose= false;
  var force= false;
  var success= false;
  
  const nextUpdateTime= GetValueByName(sheetID, "IndexStranglesBoxesUpdateNext", verbose);
  const currentTime= new Date();

  if (backupRun == undefined)
  {
    force= true;
  }
  
  if (force || (backupRun && currentTime > nextUpdateTime))
  {
    // Looks like we are due for an update
    const expirationTargets= GetTableByNameSimple(sheetID, "IndexStranglesBoxesDTEs", verbose);
    const underlying= GetValueByName(sheetID, "IndexStranglesBoxesUnderlying", verbose);
    var dteEarliest= 1000000;
    var dteLatest= 0;
    var expirations= null;

    if (expirationTargets && underlying)
    {
      // We have a list of target expirations
      for (var index in expirationTargets)
      {
        // Find expiration targets endpoints
        if (typeof expirationTargets[index][0] == "number")
        {
          if (dteLatest < expirationTargets[index][0])
          {
            dteLatest= expirationTargets[index][0]
          }
          if (dteEarliest > expirationTargets[index][0])
          {
            dteEarliest= expirationTargets[index][0]
          }
        }
      }

      if (dteEarliest <= dteLatest)
      {
        // Get the chain
        response= GetChainForSymbolByExpirationTDA(sheetID, underlying, dteEarliest, dteLatest, verbose);
      }
      else
      {
        Logger.log("[RunBoxTradeCandidates] Improper DTE endpoints: Earliest= %s, Latest= %s",
                    dteEarliest.toFixed(0), dteLatest.toFixed(0));
      }

      if (response)
      {
        // Data fetched -- extract
        expirations= ExtractExpirationsTDA(response, expirationTargets, verbose);
      }
      else
      {
        // Failed to fetch results 
        Logger.log("[RunBoxTradeCandidates] Could not fetch option chain for <%s> for expirations between <%s> and <%s> days!",
                    underlying, dteEarliest.toFixed(0), dteLatest.toFixed(0));
      }

      if (expirations)
      {
        // We have viable expiration dates -- save them
        if (SetTableByName(sheetID, "IndexStranglesBoxesExpirations", expirations, verbose))
        {
          success= true;
          UpdateTime(sheetID, "IndexStranglesBoxesUpdateTime", verbose);
          SetValueByName(sheetID, "IndexStranglesBoxesUpdateStatus", "Updated [" + DateToLocaleString(currentTime) + "]", verbose);
        }
      }
    }
    else
    {
      Logger.log("[RunBoxTradeCandidates] Missing parameters: Underlying= %s, Expiration Targets= %s", underlying, expirationTargets);
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
function FindBestBoxTradeCandidates(backupRun)
{
  // Declare constants and local variables
  const sheetID= GetMainSheetID();
  const underlying= "$SPX.X";
  const dteEarliest= 30;
  const dteLatest= 360;
  const labelPuts= "putExpDateMap";
  const labelCalls= "callExpDateMap";
  const labelPutLow= "putLow";
  const labelPutHigh= "putHigh";
  const labelCallLow= "callLow";
  const labelCallHigh= "callHigh";
  const labelSymbol= "putLowSymbol";
  const amount= 1000;
  const verbose= false;
  var yields= {};

  const response= GetChainForSymbolByExpirationTDA(sheetID, underlying, dteEarliest, dteLatest, verbose);
  if (response)
  {
    // Data fetched -- extract
    const contentParsed= ExtractContentTDA(response);

    if (contentParsed)
    {
      // Find highest yielding box trade for a given invested amount

      yields= FindBoxTradeYields(contentParsed[labelPuts], contentParsed[labelCalls], amount);
      yieldsOrdered= Object.keys(yields).sort();

      for (const ytm of yieldsOrdered)
      {
        Logger.log("%s:  %s% ($%s = $%s - $%s - $%s + $%s)", yields[ytm][labelSymbol], (ytm * 100).toFixed(2),
          (yields[ytm][labelPutHigh] - yields[ytm][labelCallHigh] - yields[ytm][labelPutLow] + yields[ytm][labelCallLow]).toFixed(2),
          yields[ytm][labelPutHigh].toFixed(2), yields[ytm][labelCallHigh].toFixed(2),
          yields[ytm][labelPutLow].toFixed(2), yields[ytm][labelCallLow].toFixed(2));

        // Logger.log("%s%:  %s", (ytm * 100).toFixed(2), yields[ytm]);
      }
    }
    else
    {
      // Did we get bad content?
      Logger.log("[FindBestBoxTradeCandidates] Could not parse returned content: <%s>", contentParsed);
    }
  }
  else
  {
    // Failed to fetch results 
    Logger.log("[FindBestBoxTradeCandidates] Could not fetch option chain for <%s> for expirations between <%s> and <%s> days!",
                underlying, dteEarliest.toFixed(0), dteLatest.toFixed(0));
  }

  return yields;
};


/**
 * FindBoxTradeYields()
 *
 * Return a list of box trade contracts with a viable yield for a given invested amount
 *
 */
function FindBoxTradeYields(puts, calls, amount)
{
  // Declare constants and local variables
  const expirations= Object.keys(puts).sort();
  const labelSymbol= "symbol";
  const expirationDelimiter= ":";
  var yields= {};

  for (const expiration of expirations)
  {
    // Collect yields from contracts which statisfy our box value (amount) for each expiration (DTE)
    const dte= expiration.split(expirationDelimiter)[1];
    const contractsPuts= puts[expiration];
    const contractsCalls= calls[expiration];
    const strikes= Object.keys(contractsPuts).sort();
    var strikeHigh= null;

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
          const pricePutLow= ExtractPriceTDA(contractsPuts[strikeLow], amount);
          const pricePutHigh= ExtractPriceTDA(contractsPuts[strikeHigh], amount);
          const priceCallLow= ExtractPriceTDA(contractsCalls[strikeLow], amount);
          const priceCallHigh= ExtractPriceTDA(contractsCalls[strikeHigh], amount);
          const yieldToMaturity= BoxYield(pricePutLow, pricePutHigh, priceCallLow, priceCallHigh, dte, amount);

          if (yieldToMaturity)
          {
            // We have a viable yield, add it to our list
            
            yields[yieldToMaturity]=
            {
              putLowSymbol : contractsPuts[strikeLow][0][labelSymbol], 
              putLow : pricePutLow,
              putHigh : pricePutHigh,
              callLow : priceCallLow,
              callHigh : priceCallHigh
            };
          }
        }
      }
    }
  }

  return yields;
};


/**
 * ExtractPriceTDA()
 *
 * Extract price from a given list of contracts for a specific strike and expiration
 *
 */
function ExtractPriceTDA(quotes, amount)
{
  // Declare constants and local variables
  const labelSymbol= "symbol";
  const labelLastPrice= "last";
  const labelClosePrice= "closePrice";
  const labelBidPrice= "bid";
  const labelAskPrice= "ask";
  const labelDelta= "delta";
  const deltaBad= "-999.0";
  const symbolDelimiter= "_";
  const weekly= "W";
  var price= null;

  for (quote of quotes)
  {
    if (quote[labelSymbol].split(symbolDelimiter)[0].endsWith(weekly) && quote[labelDelta] != deltaBad)
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
          price= last;
        }
        else if (close > 0 && close > bid && close < ask)
        { 
          // Close price seems valid
          price= close;
        }
        else
        {
          // Use the mid point of the spread for price
          price= bid + (ask - bid) / 2;
        }
      }

      if (price > amount)
      {
        // Avoid extreme strikes
        price= null;
      }

      break;
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
function BoxYield(pricePutLow, pricePutHigh, priceCallLow, priceCallHigh, dte, amount)
{
  // Declare constants and local variables
  const daysPerYear= 360;
  var yieldToMaturity= null;

  if (pricePutLow && pricePutHigh && priceCallLow && priceCallHigh)
  {
    // We appear to have real prices -- compute yield
    const cost= pricePutHigh - priceCallHigh - pricePutLow + priceCallLow;

    yieldToMaturity= Math.pow(amount / cost, daysPerYear / dte) - 1;

    if (isNaN(yieldToMaturity) || yieldToMaturity < 0)
    {
      yieldToMaturity= null;
    }
  }

  return yieldToMaturity;
};