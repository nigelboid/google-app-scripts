/**
 * Main entry point for the script
 *
 * Saves a bit of data for historical comparisons
 */
function RunPersonal(backupRun)
{
  // Declare constants and local variables
  const mainSheetID = GetMainSheetID();
  const now = new Date();
  const verbose = false;
  const verboseChanges = true;
  const confirmNumbers = true;
  
  if (mainSheetID != undefined)
  {
    // Process annual sheets
    const annualSheetIDs = GetAnnualSheetIDs(mainSheetID, verbose);
    if (annualSheetIDs != undefined)
    {
      UpdateCurrentAnnualSheet(annualSheetIDs, now, verbose, backupRun, confirmNumbers);
      SynchronizeWithAnnualSheets(mainSheetID, annualSheetIDs, now, verbose, verboseChanges);

      if (backupRun)
      {
        if (!MaintainHistoriesAnnual(annualSheetIDs[now.getFullYear()], verbose))
        {
          Log("Failed to maintain annual document histories!");
        }
      }
    }
    else
    {
      Log("Failed to obtain annual sheet IDs!");
    }
    
    UpdateMainSheet(mainSheetID, now, verbose, backupRun, confirmNumbers);
  }
  else
  {
    Log("Failed to obtain main sheet ID!");
  }
  
  // We are done -- send the accumulated log...
  LogSend(mainSheetID);
};


/**
 * RunPersonalHourly()
 *
 * Synchronize and rectify various bits current Investment documents
 */
function RunPersonalHourly()
{
  // Declare constants and local variables
  const mainSheetID = GetMainSheetID();
  const now = new Date();
  const verbose = false;
  const verboseChanges = true;
  const confirmNumbers = true;
  
  
  if (mainSheetID != undefined)
  {
    // Process annual sheets
    const annualSheetIDs = GetAnnualSheetIDs(mainSheetID, verbose);
    if (annualSheetIDs != undefined)
    {
      SynchronizeMainAndAnnualSheets(mainSheetID, annualSheetIDs, now, verbose);
      RectifyAnnualDocument(annualSheetIDs, verbose, verboseChanges);
    }
    else
    {
      Log("Failed to obtain annual sheet IDs!");
    }
    
    // process main sheet
    if (!ReconcilePortfolioHistory(mainSheetID, verbose))
    {
      Log("Failed to reconcile portfolio history!");
    }

    if (!MaintainHistoriesMain(mainSheetID, verbose))
    {
      Log("Failed to maintain main document histories!");
    }
  }
  else
  {
    Log("Failed to obtain main sheet ID!");
  }
  
  // We are done with the main sheets -- send the accumulated log...
  LogSend(mainSheetID);
  

  const allocationSheetID = GetAllocationSheetID();
  if (allocationSheetID != undefined)
  {
    // update the Allocation sheet
    UpdateAllocationSheet(allocationSheetID, now, verbose, confirmNumbers);
  }
  else
  {
    Log("Failed to obtain Allocation sheet ID!");
  }
  
  // We are done -- send the accumulated log...
  LogSend(allocationSheetID);
};


/**
 * RunAuxiliary()
 *
 * Synchronize daily records (Xfinity)
 */
function RunAuxiliary()
{
  // Declare constants and local variables
  const auxiliarySheetID = "1N3VC1SLZFBkt0jYobldHWmTRH2Zt6uUi6xZkU_yuwYM";
  const updateStampName = "UsageUpdateStamp";
  const now = new Date();
  const verbose = false;

  const updateStamp = GetValueByName(auxiliarySheetID, updateStampName, verbose);
  const updateStampDate = new Date(updateStamp);
  
  if (updateStampDate.getDate() != now.getDate())
  {
    // Synchnize today's usage to yesterday in order to start new daily differentials
    SaveValue(auxiliarySheetID, "UsageToday", "UsageYesterday", verbose);
    UpdateTime(auxiliarySheetID, updateStampName, verbose)
  }
  
  // We are done -- send the accumulated log...
  LogSend(auxiliarySheetID);
};


/**
 * UpdateMainSheet()
 *
 * Update history and save important values in the main sheet
 */
function UpdateMainSheet(mainSheetID, now, verbose, backupRun, confirmNumbers)
{
  // Declare constants and local variables
  const updateRun = false;
  
  // Preserve current net asset values
  const navHistorySheetName = "H: NAV";
  const navNames =
  {
    "FreeCashFlowTodayPositive" : null,
    "FreeCashFlowTodayNegative" : null,
    "FreeCashFlowWeek" : -10000000,
    "FreeCashFlowWeekPercent" : -1,
    "HighWaterMark" : 0,
    "AssetsNAVCurrent" : 0
  };
  SaveValuesInHistory(mainSheetID, navHistorySheetName, navNames, now, backupRun, updateRun, verbose);
  
  // Preserve some current prices for later comparisons
  SaveValue(mainSheetID, "Prices", "PricesSaved", verbose);
  
  // Preserve current cash position for later comparisons
  SaveValue(mainSheetID, "ManagedCash", "ManagedCashSaved", verbose, confirmNumbers, -2000000);
  SaveValue(mainSheetID, "SavingsCash", "SavingsCashSaved", verbose, confirmNumbers, -2000000);
  SaveValue(mainSheetID, "BondsCash", "BondsCashSaved", verbose, confirmNumbers, -2000000);
  
  // Preserve current option prices for later comparisons
  SaveValue(mainSheetID, "OptionsDerived", "OptionsSaved", verbose);
  
  // Preserve current year 401K primary and catch-up contribution amounts and dates
  SaveValue(mainSheetID, "Contribution401KAmountPrimaryThisYear", "Contribution401KAmountPrimaryThisYearSaved", verbose);
  SaveValue(mainSheetID, "Contribution401KAmountCatchUpThisYear", "Contribution401KAmountCatchUpThisYearSaved", verbose);
  SaveValue(mainSheetID, "Contribution401KDateNext", "Contribution401KDateNextSaved", verbose);
  SaveValue(mainSheetID, "Contribution401KDateLast", "Contribution401KDateLastSaved", verbose);
  
  // Preserve quarterly benchmarks
  var maxLossLimit = -1;
  if (SaveValue(mainSheetID, "BenchmarksQuarterlyThisYear", "BenchmarksQuarterlyThisYearSaved", verbose, false, maxLossLimit))
  {
    // Values updated -- update time stamp
    UpdateTime(mainSheetID, "BenchmarksQuarterlyThisYearSavedUpdateTime", verbose);
  }
};


/**
 * UpdateCurrentAnnualSheet()
 *
 * Update history and save important values in the current annual sheet
 */
function UpdateCurrentAnnualSheet(annualSheetIDs, now, verbose, backupRun, confirmNumbers)
{
  // Declare constants and local variables
  const currentYear = now.getFullYear();
  const sheetID = annualSheetIDs[currentYear];
  const updateRun = false;
  
  if (sheetID)
  {
    // Preserve current managed portfolio value
    const managedHistorySheetName = "H: mV";
    const managedNames = { "ValueManaged" : 0 };
    SaveValuesInHistory(sheetID, managedHistorySheetName, managedNames, now, backupRun, updateRun, verbose);
    
    // Preserve current nominal performance values
    var nominalHistorySheetName = "H: $";
    var nominalNames =
    {
      "AccountGains" : -2000000,
      "Gain" : -2000000,
      "IncomeFCF" : -2000000,
      "ArtValue" : 20000,
      "LenValue" : 20000,
      "GainYearOverYear" : -2000000
    };
    SaveValuesInHistory(sheetID, nominalHistorySheetName, nominalNames, now, backupRun, updateRun, verbose);
    
    // Preserve current performance percentages
    var percentHistorySheetName= "H: %";
    var percentNames=
    {
      "AccountGainRates" : -1,
      "GainRateLeveraged" : -1,
      "GainRateNominal" : -1,
      "IncomeFCFPercentage" : -1,
      "ArtGainRate" : -1,
      "LenGainRate" : -1
    };
    SaveValuesInHistory(sheetID, percentHistorySheetName, percentNames, now, backupRun, updateRun, verbose);
    
    // Preserve current budget values
    var budgetHistorySheetName= "H: B";
    var budgetLimits= [-1, -1, -1];
    var budgetNames=
    {
      "Budget" : -1,
      "Spent" : -1,
      "SpentAndObligations" : -1,
      "IncomeFCF" : -1000000
    };
    SaveValuesInHistory(sheetID, budgetHistorySheetName, budgetNames, now, backupRun, updateRun, verbose);
    
    // Reset high water marks
    SaveValue(sheetID, "HighWaterMark", "HighWaterMarkSaved", verbose, confirmNumbers, 0);
    SaveValue(sheetID, "ExpensesHigh", "ExpensesHighSaved", verbose, confirmNumbers, -1000000);
    
    // Preserve current managed portfolio values
    if (SaveValue(sheetID, "ReconcileHeld", "ReconcileHeldSaved", verbose))
    {
      UpdateTime(sheetID, "ReconcileHeldUpdateTime", verbose)
    }
  }
  else
  {
    Log("[UpdateCurrentAnnualSheet] Failed to obtain ID of the sheet for the current year <%s>!", currentYear.toFixed(0));
  }
};


/**
 * SynchronizeMainAndAnnualSheets()
 *
 * Synchronize computed values from current annual sheet to the main sheet
 */
function SynchronizeMainAndAnnualSheets(mainSheetID, annualSheetIDs, now, verbose)
{
  // declare constants and local variables
  var currentYear= now.getFullYear();
  var sourceNames= [];
  var destinationNames= [];
  var numbersOnly= false;
  
  
  if (annualSheetIDs[currentYear])
  {
    // synchronize values between the current annual sheet and the main sheet
    //   TOC
    sourceNames= ["vNemesis", "vIndex", "vSustainability", "vBreakEven", "vGain", "vGainCumulativePrior",
                  "vNemesisRunning", "vIndexRunning", "vSustainabilityRunning", "vBreakEvenRunning",
                  "HighWaterMark", "IndexStranglesLeveragePut", "IndexStranglesLeverageCall"];
    destinationNames= sourceNames;
    if (!Synchronize(annualSheetIDs[currentYear], mainSheetID, sourceNames, destinationNames, verbose, verbose, numbersOnly))
    {
      Log("[SynchronizeMainAndAnnualSheets] Failed to synchronize current annual sheet <%s> to the main sheet for range <%s>.",
                  currentYear, destinationNames);
    }
    
    //   Retirement
    sourceNames= ["IncomeGrossTotal", "IncomeNetTotal", "IncomeTrading", "GainsRealized", "GainsRealizedNet", "Expenses", "TaxesDue"];
    destinationNames= sourceNames;
    if (!Synchronize(annualSheetIDs[currentYear], mainSheetID, sourceNames, destinationNames, verbose, verbose, numbersOnly))
    {
      Log("[SynchronizeMainAndAnnualSheets] Failed to synchronize current annual sheet <%s> to the main sheet for range <%s>.",
                  currentYear, destinationNames);
    }
    
    //   Allocations
    sourceNames= ["ReturnManaged", "ReturnManagedQ1", "ReturnManagedQ2", "ReturnManagedQ3", "ReturnManagedQ4", "ReturnIndex",
                  "ReturnNemesis", "IncomeWheelPercentageWeekly", "IncomeXPercentageWeekly", "IncomeBondsPercentageWeekly", "TaxesDue"];
      
    destinationNames= sourceNames;
    if (Synchronize(annualSheetIDs[currentYear], mainSheetID, sourceNames, destinationNames, verbose, verbose))
    {
      // Values synchronized -- update time stamp
      UpdateTime(mainSheetID, "BenchmarksThisYearUpdateTime", verbose);
    }
    else
    {
      Log("[SynchronizeMainAndAnnualSheets] Failed to synchronize allocations?");
    }
  }
  else
  {
    Log("[SynchronizeMainAndAnnualSheets] Failed to obtain ID of the sheet for the current year <%s>!", currentYear.toFixed(0));
  }
};


/**
 * SynchronizeWithAnnualSheets()
 *
 * Synchronize computed values from annual sheets to the current annual sheet and to the main sheet
 */
function SynchronizeWithAnnualSheets(mainSheetID, annualSheetIDs, now, verbose, verboseChanges)
{
  // declare constants and local variables
  var currentYear= now.getFullYear();
  var priorYear= currentYear - 1;
  var earliestYear= priorYear - 3;
  var year= null;
  var sourceNames= [];
  var destinationNamesIndexed= [];
  var destinationNames= [];
  var numbersOnly= true;
  
  
  // synchronize values between legacy annual sheets and the main sheet
  sourceNames= ["vNemesis", "vIndex", "vSustainability", "vBreakEven", "vGain"];
  for (var yearIndex= 0; priorYear - yearIndex >= earliestYear; )
  {
    // create custom names for each prior year
    year= (priorYear - yearIndex).toFixed(0);
    yearIndex++;
      
    if (annualSheetIDs[year])
    {
      destinationNamesIndexed= [];
      for (const name of sourceNames)
      {
        destinationNamesIndexed.push(name + yearIndex.toFixed(0));
      }
      
      if (!Synchronize(annualSheetIDs[year], mainSheetID, sourceNames, destinationNamesIndexed, verbose, verboseChanges, numbersOnly))
      {
        Log("[SynchronizeWithAnnualSheets] Failed to synchronize annual sheet for year <%s> to the main sheet for range <%s>.",
                    year, destinationNamesIndexed);
      }
    }
    else
    {
      Log("[SynchronizeWithAnnualSheets] Failed to obtain ID of the sheet for a previous year <%s>!", year);
    }
  }
  
  
  if (annualSheetIDs[currentYear])
  {
    // synchronize values between annual sheets
    if (annualSheetIDs[priorYear])
    {
      // we have prior year, proceed
      sourceNames= ["RunRate", "vNemesisRunning", "vIndexRunning", "vSustainabilityRunning", "vBreakEvenRunning", "vGain",
                    "vGainCumulative","TaxesFederalDue", "TaxesStateDue", "TaxesFederalActualRunning", "TaxesStateActualRunning",
                    "TaxesFederalDeductions", "TaxesStateDeductions", "TaxesFederalOverpayment", "TaxesStateOverpayment",
                    "Gains1256Unrealized"];
      destinationNames= MapNames(sourceNames, "Prior", "Running");
      
      if (!Synchronize(annualSheetIDs[priorYear], annualSheetIDs[currentYear], sourceNames, destinationNames, verbose, verboseChanges,
                        numbersOnly))
      {
        Log("[SynchronizeWithAnnualSheets] Failed to synchronize annual sheet for prior year <%s> " +
                    "to the main sheet for range <%s>.", priorYear, destinationNames);
      }
    
      // Allocations Prior
      sourceNames= ["ReturnManaged", "ReturnManagedQ1", "ReturnManagedQ2", "ReturnManagedQ3", "ReturnManagedQ4", "ReturnIndex",
                    "ReturnNemesis", "IncomeWheelPercentageWeekly", "IncomeXPercentageWeekly", "IncomeBondsPercentageWeekly", "TaxesDue"];
      
      destinationNames= MapNames(sourceNames, "Prior");
      
      if (Synchronize(annualSheetIDs[priorYear], mainSheetID, sourceNames, destinationNames, verbose, verboseChanges, numbersOnly))
      {
        // Values synchronized -- update time stamp
        UpdateTime(mainSheetID, "BenchmarksLastYearUpdateTime", verbose)
      }
      else
      {
        Log("[SynchronizeWithAnnualSheets] Failed to synchronize annual sheet for prior year <%s> to the main sheet for range <%s>.",
          priorYear, destinationNames);
      }
    }
    else
    {
      Log("[SynchronizeWithAnnualSheets] Failed to obtain ID of the sheet for the prior year <%s>!", priorYear.toFixed(0));
    }
  }
  else
  {
    Log("[SynchronizeWithAnnualSheets] Failed to obtain ID of the sheet for the current year <%s>!", currentYear.toFixed(0));
  }
};


/**
 * RectifyAnnualDocument()
 *
 * Find and update stale and incorrect values in current annual document
 */
function RectifyAnnualDocument(annualSheetIDs, verbose, verboseChanges)
{
  // declare constants and local variables
  var currentYear= new Date().getFullYear();
  var previousYear= currentYear - 1;
  var accountInfoNames= ["AccountInfo", "AccountInfoKids"];
  var accountInfo= [];
  
  for (var typeIndex= 0; typeIndex < accountInfoNames.length; typeIndex++)
  {
    // check each account
    accountInfo= GetTableByNameSimple(annualSheetIDs[currentYear], accountInfoNames[typeIndex], verbose);
    
    for (var accountIndex= 0; accountIndex < accountInfo.length; accountIndex++)
    {
      if (accountInfo[accountIndex][0].length > 0)
      {
        UpdateAccountBases(annualSheetIDs[currentYear], annualSheetIDs[previousYear], accountInfo[accountIndex][0],
                            verbose, verboseChanges);
      }
    }
  }
};


/**
 * MaintainHistoriesMain()
 *
 * Sort history entries and hide older entries
 */
function MaintainHistoriesMain(mainSheetID, verbose)
{
  // declare constants and local variables
  const sheetNamesList= GetValueByName(mainSheetID, "ParametersHistorySortNames", verbose);
  const historySpecificationRange= GetValueByName(mainSheetID, "ParametersHistorySortRange", verbose);
  const columnDate= GetValueByName(mainSheetID, "ParameterPortfolioHistoryColumnDate", verbose);
  const isAscending= true;

  var spreadsheet= null;
  var rangeSpecification= "";
  var sheetNames= [];
  var range= null;
  var table= null;
  var success= true;

  
  if (sheetNamesList.length > 0)
  {
    // We seem to have entries, proceed
    sheetNames= sheetNamesList.split(",");

    for (var sheetName of sheetNames)
    {
      // Maintain each sheet listed
      sheetName= sheetName.trim();
      rangeSpecification= GetCellValue(mainSheetID, sheetName, historySpecificationRange, verbose);
      rangeSpecification= sheetName.concat("!", rangeSpecification);

      if (spreadsheet= SpreadsheetApp.openById(mainSheetID))
      {
        if (range= spreadsheet.getRange(rangeSpecification))
        {
          // check for sorting issues
          if (table= FindHistorySortFault(range, columnDate, isAscending, verbose))
          {
            // faults found -- fix them
            range= FixHistorySortFault(mainSheetID, range, table, columnDate, verbose);
            if (!range)
            {
              success= false;
              Log("[MaintainHistoriesMain] Failed to sort history!");
            }
          }
          else if (verbose)
          {
            Log("[MaintainHistoriesMain] History appears properly sorted for sheet <%s>.", sheetName);
          }
          
          // fix visibility issues, if any
          if (!FixHistoryVisibilityFault(range, columnDate, verbose))
          {
            success= false;
            Log("[MaintainHistoriesMain] Failed to adjust history visibility!");
          }
        }
        else
        {
          success= false;
          Log("[MaintainHistoriesMain] Could not get range <%s> of spreadsheet <%s>.", rangeSpecification, spreadsheet.getName());
        }
      }
      else
      {
        success= false;
        Log("[MaintainHistoriesMain] Could not open spreadsheet ID <%s>.", mainSheetID);
      }
    }
  }
  else
  {
    if (verbose)
    {
      Log("[MaintainHistoriesMain] Nothing listed for maintenance.");
    }
  }

  return success;
};


/**
 * MaintainHistoriesAnnual()
 *
 * Sort history entries and hide older entries
 */
function MaintainHistoriesAnnual(sheetID, verbose)
{
  // declare constants and local variables
  const rangeNameIncome= "Income";
  const columnDate= GetValueByName(sheetID, "ParameterIncomeSortColumn", verbose);
  const isAscending= false;
  
  var spreadsheet= null;
  var range= null;
  var success= true;
  
  if (spreadsheet= SpreadsheetApp.openById(sheetID))
  {
    if (range= TrimHistoryRange(spreadsheet.getRangeByName(rangeNameIncome), verbose))
    {
      // check for sorting issues
      if (FindHistorySortFault(range, columnDate, isAscending, verbose))
      {
        if (verbose)
        {
          Log("[MaintainHistoriesAnnual] Sort required for range named <%s> of spreadsheet <%s>.",
                      rangeNameIncome, spreadsheet.getName());

          Log("[MaintainHistoriesAnnual] Range: <%s>", range.getA1Notation());
          Log("[MaintainHistoriesAnnual] Sort column: <%s>", (columnDate + range.getColumn() - 1).toFixed(0));
          Log("[MaintainHistoriesAnnual] Sort in ascending order: <%s>", isAscending);
        }

        if (range= range.sort({column: (columnDate + range.getColumn() - 1), ascending: isAscending}))
        {
          // flush any changes applied
          SpreadsheetApp.flush();
          
          if (verbose)
          {
            Log("[MaintainHistoriesAnnual] Updated and sorted history in range <%s>.", range.getA1Notation());
          }
        }
        else
        {
          Log("[MaintainHistoriesAnnual] Failed to sort history in range <%s>.", range.getA1Notation());
        }
      }
    }
    else
    {
      success= false;
      Log("[MaintainHistoriesAnnual] Could not get range named <%s> of spreadsheet <%s>.",
                  rangeNameIncome, spreadsheet.getName());
    }
  }
  else
  {
    success= false;
    Log("[MaintainHistoriesAnnual] Could not open spreadsheet ID <%s>.", sheetID);
  }

  return success;
};

  
/**
 * ReconcilePortfolioHistory()
 *
 * Find, group, and shift redundant history entries
 */
function ReconcilePortfolioHistory(sheetID, verbose)
{
  // declare constants and local variables
  const rangeSpecification= GetValueByName(sheetID, "ParameterPortfolioHistoryRange", verbose);
  const columnDate= GetValueByName(sheetID, "ParameterPortfolioHistoryColumnDate", verbose);
  const isAscending= true;
  
  var spreadsheet= null;
  var range= null;
  var table= null;
  var success= true;
  
  
  if (spreadsheet= SpreadsheetApp.openById(sheetID))
  {
    if (range= spreadsheet.getRange(rangeSpecification))
    {
      // check for sorting issues
      if (table= FindHistorySortFault(range, columnDate, isAscending, verbose))
      {
        // faults found -- fix them
        range= FixHistorySortFault(sheetID, range, table, columnDate, verbose);
        if (!range)
        {
          success= false;
          Log("[ReconcilePortfolioHistory] Failed to sort history!");
        }
      }
      else if (verbose)
      {
        Log("[ReconcilePortfolioHistory] History appears properly sorted.");
      }
      
      // fix visibility issues, if any
      if (!FixHistoryVisibilityFault(range, columnDate, verbose))
      {
        success= false;
        Log("[ReconcilePortfolioHistory] Failed to adjust history visibility!");
      }
    }
    else
    {
      success= false;
      Log("[ReconcilePortfolioHistory] Could not get range <%s> of spreadsheet <%s>.",
                  rangeSpecification, spreadsheet.getName());
    }
  }
  else
  {
    success= false;
    Log("[ReconcilePortfolioHistory] Could not open spreadsheet ID <%s>.", sheetID);
  }

  return success;
};


/**
 * FindHistorySortFault()
 *
 * Check a history table for sorting problems
 */
function FindHistorySortFault(range, columnDateGoogle, isAscending, verbose)
{
  // declare constants and local variables
  var table= null;
  var row= null;
  var sortRequired= false;
  
  // create interesting dates
  var lastDate= null;
  var oneWeekAgo= new Date();
  oneWeekAgo.setDate(oneWeekAgo.getDate() - 7);

  // adjust column index to accommodate zero-first instead of one-first
  const columnDate= columnDateGoogle - 1;
  
  if (table= GetTableByRangeSimple(range, verbose))
  {
    // seed the last date
    lastDate= table[table.length - 1][columnDate];
    
    // check each row from the bottom to the first hidden row
    for (row= table.length - 2; row >= 0; row--)
    {
      // check each date for proper order, unless we already found such
      if ((lastDate < table[row][columnDate] && isAscending) || (lastDate > table[row][columnDate] && !isAscending))
      {
        sortRequired= true;
        if (verbose)
        {
          Log("[FindHistorySortFault] Found dates out of order at rows <%s> and <%s>:",
                      (row + range.getRow()).toFixed(0), (row + range.getRow() + 1).toFixed(0));
          Log("[FindHistorySortFault] Row <%s>: %s", (row + range.getRow()).toFixed(0), table[row]);
          Log("[FindHistorySortFault] Row <%s>: %s", (row + range.getRow() + 1).toFixed(0), table[row + 1]);
        }
        
        break;
      }

      lastDate= table[row][columnDate];
      // Log("[FindHistorySortFault] New date to compare: %s", lastDate);
    }
  }
  else
  {
    Log("[FindHistorySortFault] Could not get history table!");
  }
  
  if (sortRequired)
  {
    return table;
  }
  else
  {
    return sortRequired;
  }
};


/**
 * FixHistorySortFault()
 *
 * Adjust complementary (redundant) entries and sort history
 */
function FixHistorySortFault(sheetID, range, table, columnDate, verbose)
{
  // sorting with hidden rows triggers headaches
  range.getSheet().showRows(range.getRow(), range.getHeight());

  // since we found rows out of order, also find and match complementary history entries
  if (!UpdateComplementaryHistoryEntries(sheetID, range.getSheet().getName(), table, columnDate, range.getRow(), verbose))
  {
    // something went wrong -- report and proceed regardless
    Log("[FixHistorySortFault] Failed to update complementary history entries -- will proceed regardless...");
  }
  
  if (range= range.sort(columnDate + range.getColumn() - 1))
  {
    // flush any changes applied
    SpreadsheetApp.flush();
    
    if (verbose)
    {
      Log("[FixHistorySortFault] Updated and sorted history in range <%s>.", range.getA1Notation());
    }
  }
  else
  {
    Log("[FixHistorySortFault] Failed to sort history in range <%s>.", range.getA1Notation());
  }
  
  return range;
};


/**
 * UpdateComplementaryHistoryEntries()
 *
 * Find and update complementary history entries
 */
function UpdateComplementaryHistoryEntries(sheetID, sheetName, table, columnDateGoogle, historyOffset, verbose)
{
  // declare constants and local variables
  const columnAction= GetValueByName(sheetID, "ParameterPortfolioHistoryColumnAction", verbose) - 1;
  const columnAmount= GetValueByName(sheetID, "ParameterPortfolioHistoryColumnAmount", verbose) - 1;
  const actionCash= "Cash";
  
  var cell= null;
  var cashEntries= {};
  var columnDateSpecification= null;
  var cellCoordinates= "";
  var success= true;
  var tomorrow= new Date();
  tomorrow.setDate(tomorrow.getDate() + 1);

  
  // adjust column index to accommodate zero-first instead of one-first
  const columnDate= columnDateGoogle - 1;

  if (columnDateSpecification= SpecifyColumnA1Notation(columnDate, verbose))
  {
    // we have a viable sort column, proceed
    for (var row= 0; row < table.length; row++)
    {
      // check for complementary cash entries
      if (table[row][columnAction] == actionCash)
      {
        if (cashEntries[table[row][columnDate]] == undefined)
        {
          // unique entry (date) -- remember it
          cashEntries[table[row][columnDate]]= {};
          cashEntries[table[row][columnDate]][table[row][columnAmount]]= row;
        }
        else if (cashEntries[table[row][columnDate]][-table[row][columnAmount]] == undefined)
        {
          // unique entry (amount) -- remember it
          cashEntries[table[row][columnDate]][table[row][columnAmount]]= row;
        }
        else
        {
          // found a matching complement
          if (verbose)
          {
            Log("[UpdateComplementaryHistoryEntries] Found matching transactions on <%s> for $<%s> at rows <%s> and <%s>",
                       table[row][columnDate], table[row][columnAmount],
                       (cashEntries[table[row][columnDate]][-table[row][columnAmount]] + historyOffset).toFixed(0),
                       (row + historyOffset).toFixed(0));
          }
          
          // update dates to tomorrow prior to sorting (sort will push them to the bottom)
          cellCoordinates= columnDateSpecification + (row + historyOffset);
          if (cell= SetCellValue(sheetID, sheetName, cellCoordinates, tomorrow, verbose))
          {
            cellCoordinates= columnDateSpecification + (cashEntries[table[row][columnDate]][-table[row][columnAmount]] + historyOffset);
            if (cell= SetCellValue(sheetID, sheetName, cellCoordinates, tomorrow, verbose))
            {
              cashEntries[table[row][columnDate]][-table[row][columnAmount]]= undefined;
            }
            else
            {
              Log("[UpdateComplementaryHistoryEntries] Could not update date to <%s> for cell <%s> in sheet named <%s>.",
                          tomorrow, cellCoordinates, sheetName);
              success= false;
            }
          }
          else
          {
            Log("[UpdateComplementaryHistoryEntries] Could not update date to <%s> for cell <%s> in sheet named <%s>.",
                        tomorrow, cellCoordinates, sheetName);
            success= false;
          }
        }
      }
    }
  }
  else
  {
    Log("[UpdateComplementaryHistoryEntries] No support (yet) for managing history at column <%s>", columnDate);
    success= false;
  }
  
  // flush any changes applied
  SpreadsheetApp.flush();
  
  return success;
};


/**
 * FixHistoryVisibilityFault()
 *
 * Fix improperly shown or hidden rows (expects ascending sorted history)
 */
function FixHistoryVisibilityFault(range, columnDateGoogle, verbose)
{
  // declare constants and local variables
  var table= null;
  var row= null;
  var rowsToLowestOld= null;
  var success= true;
  var oneWeekAgo= new Date();
  
  // adjust column index to accommodate zero-first instead of one-first
  const columnDate= columnDateGoogle - 1;
  
  oneWeekAgo.setDate(oneWeekAgo.getDate() - 7);
  
  if (table= GetTableByRangeSimple(range, verbose))
  {
    // check history starting with the newest entries until encountering a date too old to show
    for (row= table.length - 1; row > 0; row--)
    {
      if (table[row][columnDate] && table[row][columnDate] < oneWeekAgo)
      {
        // found the first row too old to show or consider -- that is all we want
        rowsToLowestOld= row;
        if (verbose)
        {
          Log("[FixHistoryVisibilityFault] Found newest old row to hide in sheet <%s>: <%s>!",
                      range.getSheet().getName(), (rowsToLowestOld + range.getRow()).toFixed(0));
          Log("[FixHistoryVisibilityFault] Data: %s", table[row]);
        }
        
        break;
      }
    }
    
    // now, adjust visibility if necessary
    if (rowsToLowestOld > 1)
    {
      // hide recent data rows, except for the top one of that range and the most recent ones
      if (rowsToLowestOld == table.length - 1)
      {
        // make sure to keep at least the very last row visible
        rowsToLowestOld--;
      }
      range.getSheet().hideRows(range.getRow() + 1, rowsToLowestOld);
    }
  }
  else
  {
    Log("[FixHistoryVisibilityFault] Could not get history table from spreadsheet!");
    success= false;
  }
  
  return success;
};


/**
 * TrimHistoryRange()
 *
 * Trim specified range by dropping hidden and blank rows
 */
function TrimHistoryRange(range, verbose)
{
  // declare constants and local variables
  var topOffset= 0;
  var rows= 0;
  var row= 0;
  var table= GetTableByRangeSimple(range, verbose);
  var sheet= range.getSheet();


  if (range)
  {
    if (sheet)
    {
      if (table)
      {
        // obtain the row position of this range in the sheet
        const topRowInSheet= range.getRow();

        // find the first region of blank or hidden rows
        for (row= 0; row < table.length; row++)
        {
          // find the first hidden or blank row
          if (table[row][0] == "" || sheet.isRowHiddenByUser(row + topRowInSheet))
          {
            break;
          }
        }

        // now, find the first visible row with values
        for (; row < table.length; row++)
        {
          // find the first hidden or blank row
          if (table[row][0] != "" && !sheet.isRowHiddenByUser(row + topRowInSheet))
          {
            topOffset= row;
            if (verbose)
            {
              Log("[TrimHistoryRange] Found new top row: <%s>", (topOffset + topRowInSheet).toFixed(0));
            }
        
            break;
          }
        }

        // following, find the next blank or hidden row
        for (; row < table.length; row++)
        {
          // find the first hidden or blank row
          if (table[row][0] == "" || sheet.isRowHiddenByUser(row + topRowInSheet))
          {
            if (verbose)
            {
              Log("[TrimHistoryRange] Found new bottom row: <%s>", (row + topRowInSheet).toFixed(0));
            }
            break;
          }
        }

        // count of visible data rows to consider
        rows= row - topOffset;

        // lastly, trim the original range to contain this first chunk of visible, non-empty rows
        range= range.offset(topOffset, 0, rows);
        if (verbose)
        {
          Log("[TrimHistoryRange] Adjusted range: <%s>", range.getA1Notation());
        }
      }
      else
      {
        Log("[TrimHistoryRange] Failed to obtain values from the specified range!");
      }
    }
    else
    {
       Log("[TrimHistoryRange] Failed to obtain parent sheet for the specified range!");
    }
  }
  else
  {
    Log("[TrimHistoryRange] Range specification fault!");
  }

  return range;
};


/**
 * FixSheetNames()
 *
 * Check and update any stale account sheet names
 */
function FixSheetNames(sheetID, sheetName, verbose, verboseChanges)
{
  // declare constants and local variables
  var parameterSheetName= "ParameterAccountSheetName";
  var destinationCell= GetValueByName(sheetID, parameterSheetName, verbose);
  var sheetNameStored= GetCellValue(sheetID, sheetName, destinationCell, verbose);
  
  if (sheetNameStored != sheetName)
  {
    if (SetCellValue(sheetID, sheetName, destinationCell, sheetName, verbose))
    {
      if (verboseChanges)
      {
        Log("[FixSheetNames] Updated sheet name <%s> (it was <%s>)", sheetName, sheetNameStored);
      }
    }
    else
    {
      Log("[FixSheetNames] Failed to update sheet name <%s>", sheetName);
    }
  }
  else
  {
    if (verbose)
    {
      Log("[FixSheetNames] Stored sheet name matches sheet name <%s>", sheetName);
    }
  }
};


/**
 * UpdateAccountBases()
 *
 * Check and update any stale account basis relative to closing value from previous year
 */
function UpdateAccountBases(currentYearID, priorYearID, sheetName, verbose, verboseChanges)
{
  // declare constants and local variables
  var parameterBasisName= "ParameterAccountPriorBasis";
  var parameterValueName= "ParameterAccountValue";
  var basisCell= GetValueByName(currentYearID, parameterBasisName, verbose);
  var valueCell= GetValueByName(priorYearID, parameterValueName, verbose);
  var basisCurrentYear= GetCellValue(currentYearID, sheetName, basisCell, verbose);
  var valuePriorYear= GetCellValue(priorYearID, sheetName, valueCell, verbose);
  
  if (isNaN(valuePriorYear))
  {
    Log("[UpdateAccountBases] Failed to obtain prior year basis <%s> for accont <%s>!", valuePriorYear, sheetName);
  }
  else
  {
    if (basisCurrentYear != valuePriorYear)
    {
      if (SetCellValue(currentYearID, sheetName, basisCell, valuePriorYear, verbose))
      {
        if (verboseChanges)
        {
          Log("[UpdateAccountBases] Updated prior year basis for account <%s> to <%s> (it was <%s>).",
                      sheetName, valuePriorYear, basisCurrentYear);
        }
      }
      else
      {
        Log("[UpdateAccountBases] Failed to update prior year basis for account <%s> to <%s> (it is still <%s>).",
                    sheetName, valuePriorYear, basisCurrentYear);
      }
    }
    else
    {
      if (verbose)
      {
        Log("[UpdateAccountBases] Currently stored prior year basis <%s> for account <%s> matches value <%s> from previous year.",
                   valuePriorYear, sheetName, basisCurrentYear);
      }
    }
  }
};


/**
 * UpdateAllocationSheet()
 *
 * Update history and save important values in the public allocation sheet
 */
function UpdateAllocationSheet(sheetID, now, verbose, confirmNumbers)
{
  const updateRun = true;
  const backupRun = false;
  var historySheetName = null;
  var names = null;
  
  if (sheetID)
  {
    // Preserve current performance percentages
    historySheetName = "H: %";
    names =
    {
      "PerformanceThisYear" : -0.99,
      "IndexThisYear" : -0.99,
      "NemesisThisYear" : -0.99
    };
    SaveValuesInHistory(sheetID, historySheetName, names, now, backupRun, updateRun, verbose);
    
    // Preserve RUT performance statistics
    historySheetName = "H: RUT";
    names=
    {
      "RUTPrice" : 0,
      "RUTBufferPuts" : 0,
      "RUTBufferCalls" : 0,
      "RUTProfit" : -5
    };
    SaveValuesInHistory(sheetID, historySheetName, names, now, backupRun, updateRun, verbose);
    
    // Preserve SPX performance statistics
    historySheetName = "H: SPX";
    names=
    {
      "SPXPrice" : 0,
      "SPXBufferPuts" : 0,
      "SPXBufferCalls" : 0,
      "SPXProfit" : -5
    };
    SaveValuesInHistory(sheetID, historySheetName, names, now, backupRun, updateRun, verbose);
  }
  else
  {
    Log("Failed to obtain ID of the Allocations (public) sheet!");
  }
};


/**
 * GetAllocationSheetID()
 *
 * Return the ID of the public Allocations sheet
 */
function GetAllocationSheetID()
{
  //return "1EumQvi51zI0qWVPST2TOXyA6Gbjt0qOr2klj2I8ljfg";
  return "1lB6ndpg16wT14JIOiHtKbs_sdn_ks_qFHJKP5IN6ljg";
};


/**
 * MapNames()
 *
 * Map one set of names to another by adding a suffix and optionally, dropping an existing suffix
 */
function MapNames(sourceNames, suffix, dropSuffix)
{
  var mappedNames= [];
  
  if (dropSuffix != undefined)
  {
    // map source names to new names by repalcing a current suffix with a new one
    var dropSuffixExpression= new RegExp(dropSuffix + "$", "")
    
    for (var index= 0; index < sourceNames.length; index++)
    {
      if (dropSuffixExpression.test(sourceNames[index]))
      {
        // replace an existing suffix
        mappedNames.push(sourceNames[index].replace(dropSuffixExpression, suffix));
      }
      else
      {
        // did not match an existing suffix, just add the new one
        mappedNames.push(sourceNames[index] + suffix);
      }
    }
  }
  else
  {
    // map source names to new names by adding a suffix
    for (var index= 0; index < sourceNames.length; index++)
    {
      mappedNames.push(sourceNames[index] + suffix);
    }
  }
  
  return mappedNames;
};