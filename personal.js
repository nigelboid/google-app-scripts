/**
 * Main entry point for the script
 *
 * Saves a bit of data for historical comparisons
 */
function RunPersonal(backupRun)
{
  // Declare constants and local variables
  const mainSheetID = GetMainSheetID();
  const scriptTime = new Date();
  const verbose = false;
  const verboseChanges = true;
  const confirmNumbers = true;
  
  if (mainSheetID != undefined)
  {
    // Process annual sheets
    const annualSheetIDs = GetAnnualSheetIDs(mainSheetID, verbose);
    if (annualSheetIDs != undefined)
    {
      UpdateCurrentAnnualSheet(annualSheetIDs, scriptTime, verbose, backupRun, confirmNumbers);
      SynchronizeWithAnnualSheets(mainSheetID, annualSheetIDs, scriptTime, verbose, verboseChanges);

      if (backupRun)
      {
        if (!MaintainHistoriesAnnual(annualSheetIDs[scriptTime.getFullYear()], verbose))
        {
          Log("Failed to maintain annual document histories!");
        }
      }
    }
    else
    {
      Log("Failed to obtain annual sheet IDs!");
    }
    
    UpdateMainSheet(mainSheetID, scriptTime, verbose, backupRun, confirmNumbers);
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
  const scriptTime = new Date();
  const verbose = false;
  const verboseChanges = true;
  const confirmNumbers = true;
  
  
  if (mainSheetID != undefined)
  {
    // Process annual sheets
    const annualSheetIDs = GetAnnualSheetIDs(mainSheetID, verbose);
    if (annualSheetIDs != undefined)
    {
      SynchronizeMainAndAnnualSheets(mainSheetID, annualSheetIDs, scriptTime, verbose);
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
    UpdateAllocationSheet(allocationSheetID, scriptTime, verbose, confirmNumbers);
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
  const scriptTime = new Date();
  const verbose = false;

  const updateStamp = GetValueByName(auxiliarySheetID, updateStampName, verbose);
  const updateStampDate = new Date(updateStamp);
  
  if (updateStampDate.getDate() != scriptTime.getDate())
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
function UpdateMainSheet(mainSheetID, scriptTime, verbose, backupRun, confirmNumbers)
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
  SaveValuesInHistory(mainSheetID, navHistorySheetName, navNames, scriptTime, backupRun, updateRun, verbose);
  
  // Preserve some current prices for later comparisons
  SaveValue(mainSheetID, "Prices", "PricesSaved", verbose);
  
  // Preserve current cash position for later comparisons
  const maxDebt = -2000000;
  SaveValue(mainSheetID, "ManagedCash", "ManagedCashSaved", verbose, confirmNumbers, maxDebt);
  SaveValue(mainSheetID, "SavingsCash", "SavingsCashSaved", verbose, confirmNumbers, maxDebt);
  SaveValue(mainSheetID, "BondsCash", "BondsCashSaved", verbose, confirmNumbers, maxDebt);
  
  // Preserve current option prices for later comparisons
  SaveValue(mainSheetID, "OptionsDerived", "OptionsSaved", verbose);
  
  // Preserve current year 401K primary and catch-up contribution amounts and dates
  SaveValue(mainSheetID, "Contribution401KAmountPrimaryThisYear", "Contribution401KAmountPrimaryThisYearSaved", verbose);
  SaveValue(mainSheetID, "Contribution401KAmountCatchUpThisYear", "Contribution401KAmountCatchUpThisYearSaved", verbose);
  SaveValue(mainSheetID, "Contribution401KDateNext", "Contribution401KDateNextSaved", verbose);
  SaveValue(mainSheetID, "Contribution401KDateLast", "Contribution401KDateLastSaved", verbose);
  
  // Preserve quarterly benchmarks
  const maxLossLimit = -1;
  const confirmMyNumbers = true;
  if (SaveValue(mainSheetID, "BenchmarksQuarterlyThisYear", "BenchmarksQuarterlyThisYearSaved", verbose, confirmMyNumbers, maxLossLimit))
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
function UpdateCurrentAnnualSheet(annualSheetIDs, scriptTime, verbose, backupRun, confirmNumbers)
{
  // Declare constants and local variables
  const currentYear = scriptTime.getFullYear();
  const sheetID = annualSheetIDs[currentYear];
  const updateRun = false;
  
  if (sheetID)
  {
    // Preserve current managed portfolio value
    const managedHistorySheetName = "H: mV";
    const managedNames = { "ValueManaged" : 0 };
    SaveValuesInHistory(sheetID, managedHistorySheetName, managedNames, scriptTime, backupRun, updateRun, verbose);
    
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
    SaveValuesInHistory(sheetID, nominalHistorySheetName, nominalNames, scriptTime, backupRun, updateRun, verbose);
    
    // Preserve current performance percentages
    var percentHistorySheetName = "H: %";
    var percentNames =
    {
      "AccountGainRates" : -1,
      "GainRateLeveraged" : -1,
      "GainRateNominal" : -1,
      "IncomeFCFPercentage" : -1,
      "ArtGainRate" : -1,
      "LenGainRate" : -1
    };
    SaveValuesInHistory(sheetID, percentHistorySheetName, percentNames, scriptTime, backupRun, updateRun, verbose);
    
    // Preserve current budget values
    var budgetHistorySheetName = "H: B";
    var budgetNames =
    {
      "Budget" : -1,
      "Spent" : -1,
      "SpentAndObligations" : -1,
      "IncomeFCF" : -1000000
    };
    SaveValuesInHistory(sheetID, budgetHistorySheetName, budgetNames, scriptTime, backupRun, updateRun, verbose);
    
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
    Log(`Failed to obtain ID of the sheet for the current year <${currentYear.toFixed(0)}>!`);
  }
};


/**
 * SynchronizeMainAndAnnualSheets()
 *
 * Synchronize computed values from current annual sheet to the main sheet
 */
function SynchronizeMainAndAnnualSheets(mainSheetID, annualSheetIDs, scriptTime, verbose)
{
  // Declare constants and local variables
  const currentYear = scriptTime.getFullYear();
  const numbersOnly = false;
  const destinationNames = null;
  var sourceNames = null;
  
  
  if (annualSheetIDs[currentYear])
  {
    // Synchronize values between the current annual sheet and the main sheet
    //   TOC
    sourceNames =
    [
      "vNemesis",
      "vIndex",
      "vSustainability",
      "vBreakEven",
      "vGain",
      "vGainCumulativePrior",
      "vNemesisRunning",
      "vIndexRunning",
      "vSustainabilityRunning",
      "vBreakEvenRunning",
      "HighWaterMark",
      "IndexStranglesLeveragePut",
      "IndexStranglesLeverageCall"
    ];
    if (!Synchronize(annualSheetIDs[currentYear], mainSheetID, sourceNames, destinationNames, verbose, verbose, numbersOnly))
    {
      Log(`Failed to synchronize current annual sheet <${currentYear}> to the main sheet for range <${sourceNames}>!`);
    }
    
    //   Retirement
    sourceNames =
    [
      "IncomeGrossTotal",
      "IncomeNetTotal",
      "IncomeTrading",
      "GainsRealized",
      "GainsRealizedNet",
      "Expenses",
      "TaxesDue"
    ];
    if (!Synchronize(annualSheetIDs[currentYear], mainSheetID, sourceNames, destinationNames, verbose, verbose, numbersOnly))
    {
      Log(`Failed to synchronize current annual sheet <${currentYear}> to the main sheet for range <${sourceNames}>!`);
    }
    
    //   Allocations
    sourceNames =
    [
      "ReturnManaged",
      "ReturnManagedQ1",
      "ReturnManagedQ2",
      "ReturnManagedQ3",
      "ReturnManagedQ4",
      "ReturnIndex",
      "ReturnNemesis",
      "IncomeWheelPercentageWeekly",
      "IncomeXPercentageWeekly",
      "IncomeBondsPercentageWeekly",
      "TaxesDue"
    ];
    if (Synchronize(annualSheetIDs[currentYear], mainSheetID, sourceNames, destinationNames, verbose, verbose))
    {
      // Values synchronized -- update time stamp
      UpdateTime(mainSheetID, "BenchmarksThisYearUpdateTime", verbose);
    }
    else
    {
      Log(`Failed to synchronize current annual sheet <${currentYear}> to the main sheet for range <${sourceNames}>!`);
    }
  }
  else
  {
    Log(`Failed to obtain ID of the sheet for the current year <${currentYear.toFixed(0)}>!` );
  }
};


/**
 * SynchronizeWithAnnualSheets()
 *
 * Synchronize computed values from annual sheets to the current annual sheet and to the main sheet
 */
function SynchronizeWithAnnualSheets(mainSheetID, annualSheetIDs, scriptTime, verbose, verboseChanges)
{
  // declare constants and local variables
  const currentYear = scriptTime.getFullYear();
  const priorYear = currentYear - 1;
  const earliestYear = priorYear - 3;
  const numbersOnly = true;
  var year = null;
  var yearIndex = 0;
  var sourceNames = null;
  var destinationNamesIndexed = null;
  var destinationNames = null;
  var success = false;
  
  // synchronize values between legacy annual sheets and the main sheet
  sourceNames =
  [
    "vNemesis",
    "vIndex",
    "vSustainability",
    "vBreakEven",
    "vGain"
  ];
  while (priorYear - yearIndex >= earliestYear)
  {
    // Create custom names for each prior year
    year = (priorYear - yearIndex).toFixed(0);
    yearIndex++;
      
    if (annualSheetIDs[year])
    {
      destinationNamesIndexed = [];
      for (const name of sourceNames)
      {
        destinationNamesIndexed.push(name + yearIndex.toFixed(0));
      }
      
      if (!Synchronize(annualSheetIDs[year], mainSheetID, sourceNames, destinationNamesIndexed, verbose, verboseChanges, numbersOnly))
      {
        Log(`Failed to synchronize annual sheet for year <${year}> to the main sheet for range <${destinationNamesIndexed}>.`);
      }
    }
    else
    {
      Log(`Failed to obtain ID of the sheet for a previous year <${year}>!`);
    }
  }
  
  
  if (annualSheetIDs[currentYear])
  {
    // Synchronize values between annual sheets
    if (annualSheetIDs[priorYear])
    {
      // Proceed with prior year
      sourceNames =
      [
        "RunRate",
        "vNemesisRunning",
        "vIndexRunning",
        "vSustainabilityRunning",
        "vBreakEvenRunning",
        "vGain",
        "vGainCumulative",
        "TaxesFederalDue",
        "TaxesStateDue",
        "TaxesFederalActualRunning",
        "TaxesStateActualRunning",
        "TaxesFederalDeductions",
        "TaxesStateDeductions",
        "TaxesFederalOverpayment",
        "TaxesStateOverpayment",
        "Gains1256Unrealized"
      ];
      destinationNames = MapNames(sourceNames, "Prior", "Running");
      
      success = Synchronize
      (
        annualSheetIDs[priorYear], annualSheetIDs[currentYear], sourceNames, destinationNames,
        verbose, verboseChanges, numbersOnly
      );
      if (!success)
      {
        Log(`Failed to synchronize annual sheet for prior year <${priorYear}> to the main sheet for range <${destinationNames}>.`);
      }
    
      // Allocations Prior
      sourceNames =
      [
        "ReturnManaged",
        "ReturnManagedQ1",
        "ReturnManagedQ2",
        "ReturnManagedQ3",
        "ReturnManagedQ4",
        "ReturnIndex",
        "ReturnNemesis",
        "IncomeWheelPercentageWeekly",
        "IncomeXPercentageWeekly",
        "IncomeBondsPercentageWeekly",
        "TaxesDue"
      ];
      destinationNames = MapNames(sourceNames, "Prior");
      
      if (Synchronize(annualSheetIDs[priorYear], mainSheetID, sourceNames, destinationNames, verbose, verboseChanges, numbersOnly))
      {
        // Values synchronized -- update time stamp
        UpdateTime(mainSheetID, "BenchmarksLastYearUpdateTime", verbose)
      }
      else
      {
        Log(`Failed to synchronize annual sheet for prior year <${priorYear}> to the main sheet for range <${destinationNames}>.`);
      }
    }
    else
    {
      Log(`Failed to obtain ID of the sheet for the prior year <${priorYear.toFixed(0)}>!`);
    }
  }
  else
  {
    Log(`Failed to obtain ID of the sheet for the current year <${currentYear.toFixed(0)}>!`);
  }
};


/**
 * RectifyAnnualDocument()
 *
 * Find and update stale and incorrect values in current annual document
 */
function RectifyAnnualDocument(annualSheetIDs, verbose, verboseChanges)
{
  // Declare constants and local variables
  const currentYear = new Date().getFullYear();
  const previousYear = currentYear - 1;
  const accountInfoNames = ["AccountInfo", "AccountInfoKids"];
  var accountInfo = null;
  
  for (const typeIndex in accountInfoNames)
  {
    // Check each account
    accountInfo = GetTableByNameSimple(annualSheetIDs[currentYear], accountInfoNames[typeIndex], verbose);
    
    for (const accountIndex in accountInfo)
    {
      if (accountInfo[accountIndex][0].length > 0)
      {
        UpdateAccountBases
        (
          annualSheetIDs[currentYear], annualSheetIDs[previousYear], accountInfo[accountIndex][0],
          verbose, verboseChanges
        );
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
  var rangeSpecification= null;
  var range= null;
  var table= null;
  var success= true;

  
  if (sheetNamesList.length > 0)
  {
    // We seem to have entries, proceed
    const sheetNames = sheetNamesList.split(",");

    for (var sheetName of sheetNames)
    {
      // Maintain each sheet listed
      sheetName = sheetName.trim();
      rangeSpecification = GetCellValue(mainSheetID, sheetName, historySpecificationRange, verbose);
      rangeSpecification = sheetName.concat("!", rangeSpecification);

      spreadsheet = SpreadsheetApp.openById(mainSheetID)
      if (spreadsheet)
      {
        range = spreadsheet.getRange(rangeSpecification)
        if (range)
        {
          // Check for sorting issues
          table = FindHistorySortFault(range, columnDate, isAscending, verbose)
          if (table)
          {
            // Faults found -- fix them
            range= FixHistorySortFault(mainSheetID, range, table, columnDate, verbose);
            if (!range)
            {
              success = false;
              Log("Failed to sort history!");
            }
          }
          else if (verbose)
          {
            Log(`History appears properly sorted for sheet <${sheetName}>.`);
          }
          
          // Fix visibility issues, if any
          if (!FixHistoryVisibilityFault(range, columnDate, verbose))
          {
            success = false;
            Log("Failed to adjust history visibility!");
          }
        }
        else
        {
          success = false;
          Log(`Could not get range <${rangeSpecification}> of spreadsheet <${spreadsheet.getName()}>.`);
        }
      }
      else
      {
        success = false;
        Log(`Could not open spreadsheet ID <${mainSheetID}>.`);
      }
    }
  }
  else
  {
    LogVerbose("Nothing listed for maintenance.", verbose);
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
  // Declare constants and local variables
  const rangeNameIncome = "Income";
  const columnDate = GetValueByName(sheetID, "ParameterIncomeSortColumn", verbose);
  const isAscending = false;
  const spreadsheet = SpreadsheetApp.openById(sheetID);
  var range = null;
  var success = true;
   
  if (spreadsheet)
  {
    range = TrimHistoryRange(spreadsheet.getRangeByName(rangeNameIncome), verbose);
    if (range)
    {
      // Check for sorting issues
      if (FindHistorySortFault(range, columnDate, isAscending, verbose))
      {
        LogVerbose(`Sort required for range named <${rangeNameIncome}> of spreadsheet <${spreadsheet.getName()}>.`);
        LogVerbose(`Range: <${range.getA1Notation()}>`);
        LogVerbose(`Sort column: <${(columnDate + range.getColumn() - 1).toFixed(0)}>`);
        LogVerbose(`Sort in ascending order: <${isAscending}>`);

        range = range.sort({column: (columnDate + range.getColumn() - 1), ascending: isAscending});
        if (range)
        {
          // Flush any changes applied
          SpreadsheetApp.flush();
          
          LogVerbose(`Updated and sorted history in range <${range.getA1Notation()}>.`, verbose);
        }
        else
        {
          Log(`Failed to sort history in range <${range.getA1Notation()}>.`);
        }
      }
    }
    else
    {
      success = false;
      Log(`Could not get range named <${rangeNameIncome}> of spreadsheet <${spreadsheet.getName()}>.`);
    }
  }
  else
  {
    success = false;
    Log(`Could not open spreadsheet ID <${sheetID}>.`);
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
  const rangeSpecification = GetValueByName(sheetID, "ParameterPortfolioHistoryRange", verbose);
  const columnDate = GetValueByName(sheetID, "ParameterPortfolioHistoryColumnDate", verbose);
  const isAscending = true;
  const spreadsheet = SpreadsheetApp.openById(sheetID);
  var range = null;
  var table = null;
  var success = true;
  
  
  if (spreadsheet)
  {
    range = spreadsheet.getRange(rangeSpecification);
    if (range)
    {
      // Check for sorting issues
      table = FindHistorySortFault(range, columnDate, isAscending, verbose)
      if (table)
      {
        // Faults found -- fix them
        range = FixHistorySortFault(sheetID, range, table, columnDate, verbose);
        if (!range)
        {
          success = false;
          Log("Failed to sort history!");
        }
      }
      else if (verbose)
      {
        Log("History appears properly sorted.");
      }
      
      // Fix visibility issues, if any
      if (!FixHistoryVisibilityFault(range, columnDate, verbose))
      {
        success = false;
        Log("Failed to adjust history visibility!");
      }
    }
    else
    {
      success = false;
      Log(`Could not get range <${rangeSpecification}> of spreadsheet <${spreadsheet.getName()}>.`);
    }
  }
  else
  {
    success = false;
    Log(`Could not open spreadsheet ID <${sheetID}>.`);
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
  // Declare constants and local variables
  const table = GetTableByRangeSimple(range, verbose);;
  var row = null;
  var sortRequired = false;
  
  // Create interesting dates
  var lastDate = null;
  var oneWeekAgo = new Date();
  oneWeekAgo.setDate(oneWeekAgo.getDate() - 7);

  // Adjust column index to accommodate zero-first instead of one-first
  const columnDate = columnDateGoogle - 1;
  
  if (table)
  {
    // Seed the last date
    lastDate = table[table.length - 1][columnDate];
    
    // Check each row from the bottom to the first hidden row
    for (row = table.length - 2; row >= 0; row--)
    {
      // check each date for proper order, unless we already found such
      if ((lastDate < table[row][columnDate] && isAscending) || (lastDate > table[row][columnDate] && !isAscending))
      {
        const outOfOrderRow = (row + range.getRow());
        LogVerbose(`Found dates out of order at row <${outOfOrderRow.toFixed(0)}>!`, verbose);
        LogVerbose(`Row <${outOfOrderRow.toFixed(0)}>: ${table[row]}`, verbose);
        LogVerbose(`Row <${(outOfOrderRow + 1).toFixed(0)}>: ${table[row + 1]}`, verbose);
        
        sortRequired = true;
        break;
      }

      lastDate = table[row][columnDate];
    }
  }
  else
  {
    Log("Could not get history table!");
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
  // Sorting with hidden rows triggers headaches
  range.getSheet().showRows(range.getRow(), range.getHeight());

  // Since we found rows out of order, also find and match complementary history entries
  if (!UpdateComplementaryHistoryEntries(sheetID, range.getSheet().getName(), table, columnDate, range.getRow(), verbose))
  {
    // something went wrong -- report and proceed regardless
    Log("Failed to update complementary history entries -- will proceed regardless...");
  }
  
  range = range.sort(columnDate + range.getColumn() - 1);
  if (range)
  {
    // Flush any changes applied
    SpreadsheetApp.flush();
    
    LogVerbose(`Updated and sorted history in range <${range.getA1Notation()}>.`, verbose);
  }
  else
  {
    Log(`Failed to sort history in range <${range.getA1Notation()}>.`);
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
  const columnAction = GetValueByName(sheetID, "ParameterPortfolioHistoryColumnAction", verbose) - 1;
  const columnAmount = GetValueByName(sheetID, "ParameterPortfolioHistoryColumnAmount", verbose) - 1;
  const actionCash = "Cash";
  
  var cell = null;
  var cashEntries = {};
  var columnDateSpecification = null;
  var cellCoordinates = "";
  var success = true;

  // Definie "tomorrow" as the day after the current day
  const tomorrow = new Date();
  tomorrow.setDate(tomorrow.getDate() + 1);

  // Adjust column index to accommodate zero-first instead of one-first
  const columnDate = columnDateGoogle - 1;

  columnDateSpecification = SpecifyColumnA1Notation(columnDate, verbose)
  if (columnDateSpecification)
  {
    // Found a viable sort column, proceed
    for (var row = 0; row < table.length; row++)
    {
      // Check for complementary cash entries
      if (table[row][columnAction] == actionCash)
      {
        if (cashEntries[table[row][columnDate]] == undefined)
        {
          // Unique entry (date) -- preserve it
          cashEntries[table[row][columnDate]] = {};
          cashEntries[table[row][columnDate]][table[row][columnAmount]] = row;
        }
        else if (cashEntries[table[row][columnDate]][-table[row][columnAmount]] == undefined)
        {
          // Unique entry (amount) -- preserve it
          cashEntries[table[row][columnDate]][table[row][columnAmount]] = row;
        }
        else
        {
          // Found a matching complement
          LogVerbose
          (
            `Found matching transactions on <${table[row][columnDate]}> for $<${table[row][columnAmount]}> ` +
            `at rows <${(cashEntries[table[row][columnDate]][-table[row][columnAmount]] + historyOffset)}> ` +
            `and <${(row + historyOffset)}>`,
            verbose
          );
          
          // Update dates to tomorrow prior to sorting (sort will push them to the bottom)
          cellCoordinates = columnDateSpecification + (row + historyOffset);
          cell = SetCellValue(sheetID, sheetName, cellCoordinates, tomorrow, verbose);
          if (cell)
          {
            cellCoordinates = columnDateSpecification + (cashEntries[table[row][columnDate]][-table[row][columnAmount]] + historyOffset);
            cell = SetCellValue(sheetID, sheetName, cellCoordinates, tomorrow, verbose);
            if (cell)
            {
              cashEntries[table[row][columnDate]][-table[row][columnAmount]] = undefined;
            }
            else
            {
              Log(`Could not update date to <${tomorrow}> for cell <${cellCoordinates}> in sheet named <${sheetName}>.`);
              success = false;
            }
          }
          else
          {
            Log(`Could not update date to <${tomorrow}> for cell <${cellCoordinates}> in sheet named <${sheetName}>.`);
            success = false;
          }
        }
      }
    }
  }
  else
  {
    Log(`No support (yet) for managing history at column <${columnDate}>`);
    success = false;
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
  // Declare constants and local variables
  const table = GetTableByRangeSimple(range, verbose);
  var rowsToLowestOld = null;
  var oldRow = null;
  var success = true;
  
  // Adjust column index to accommodate zero-first instead of one-first
  const columnDate= columnDateGoogle - 1;
  
  const oneWeekAgo = new Date();
  oneWeekAgo.setDate(oneWeekAgo.getDate() - 7);
  
  if (table)
  {
    // Check history starting with the newest entries until encountering a date too old to show
    for (var row = table.length - 1; row > 0; row--)
    {
      if (table[row][columnDate] && table[row][columnDate] < oneWeekAgo)
      {
        // found the first row too old to show or consider -- that is all we want
        rowsToLowestOld = row;
        oldRow = (rowsToLowestOld + range.getRow());
        LogVerbose(`Found newest old row to hide in sheet <${range.getSheet().getName()}>: <${oldRow.toFixed(0)}>!`, verbose);
        LogVerbose(`Data: ${table[row]}`, verbose);
        
        break;
      }
    }
    
    // Adjust visibility if necessary
    if (rowsToLowestOld > 1)
    {
      // Hide recent data rows, except for the top one of that range and the most recent ones
      if (rowsToLowestOld == (table.length - 1))
      {
        // Keep at least the very last row visible
        rowsToLowestOld--;
      }
      range.getSheet().hideRows(range.getRow() + 1, rowsToLowestOld);
    }
  }
  else
  {
    Log("Could not get history table from spreadsheet!");
    success = false;
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
  // Declare constants and local variables
  var topOffset = 0;
  var rows = 0;
  var row = 0;
  
  if (range)
  {
    const sheet = range.getSheet();
    if (sheet)
    {
      const table = GetTableByRangeSimple(range, verbose);
      if (table)
      {
        // Obtain the row position of this range in the sheet
        const topRowInSheet = range.getRow();

        // Find the first region of blank or hidden rows
        for (row = 0; row < table.length; row++)
        {
          // find the first hidden or blank row
          if (table[row][0] == "" || sheet.isRowHiddenByUser(row + topRowInSheet))
          {
            break;
          }
        }

        // Find the first visible row with values
        for (; row < table.length; row++)
        {
          // Find the first hidden or blank row
          if (table[row][0] != "" && !sheet.isRowHiddenByUser(row + topRowInSheet))
          {
            topOffset = row;
            LogVerbose(`Found new top row: <${(topOffset + topRowInSheet).toFixed(0)}>`, verbose);
            break;
          }
        }

        // Find the next blank or hidden row
        for (; row < table.length; row++)
        {
          // Find the first hidden or blank row
          if (table[row][0] == "" || sheet.isRowHiddenByUser(row + topRowInSheet))
          {
            LogVerbose(`Found new bottom row: <${(row + topRowInSheet).toFixed(0)}>`, verbose);
            break;
          }
        }

        // Count of visible data rows to consider
        rows = row - topOffset;

        // Trim the original range to contain this first chunk of visible, non-empty rows
        range = range.offset(topOffset, 0, rows);
        LogVerbose(`Adjusted range: <${range.getA1Notation()}>`, verbose);
      }
      else
      {
        Log("Failed to obtain values from the specified range!");
      }
    }
    else
    {
       Log("Failed to obtain parent sheet for the specified range!");
    }
  }
  else
  {
    Log("Range specification fault!");
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
  // Declare constants and local variables
  const parameterSheetName = "ParameterAccountSheetName";
  const destinationCell = GetValueByName(sheetID, parameterSheetName, verbose);
  const sheetNameStored = GetCellValue(sheetID, sheetName, destinationCell, verbose);
  
  if (sheetNameStored != sheetName)
  {
    if (SetCellValue(sheetID, sheetName, destinationCell, sheetName, verbose))
    {
      LogVerbose(`Updated sheet name <${sheetName}> (it was <${sheetNameStored}>)`, verboseChanges);
    }
    else
    {
      Log(`Failed to update sheet name <${sheetName}>`);
    }
  }
  else
  {
    LogVerbose(`Stored sheet name matches sheet name <${sheetName}>`, verbose);
  }
};


/**
 * UpdateAccountBases()
 *
 * Check and update any stale account basis relative to closing value from previous year
 */
function UpdateAccountBases(currentYearID, priorYearID, sheetName, verbose, verboseChanges)
{
  // Declare constants and local variables
  const parameterBasisName = "ParameterAccountPriorBasis";
  const parameterValueName = "ParameterAccountValue";
  const basisCell = GetValueByName(currentYearID, parameterBasisName, verbose);
  const valueCell = GetValueByName(priorYearID, parameterValueName, verbose);
  const basisCurrentYear = GetCellValue(currentYearID, sheetName, basisCell, verbose);
  const valuePriorYear = GetCellValue(priorYearID, sheetName, valueCell, verbose);
  
  if (isNaN(valuePriorYear))
  {
    Log(`Failed to obtain prior year basis <${valuePriorYear}> for accont <${sheetName}>!`);
  }
  else
  {
    if (basisCurrentYear != valuePriorYear)
    {
      if (SetCellValue(currentYearID, sheetName, basisCell, valuePriorYear, verbose))
      {
        LogVerbose
        (
          `Updated prior year basis for account <${sheetName}> to <${valuePriorYear}> (it was <${basisCurrentYear}>).`,
          verboseChanges
        );
      }
      else
      {
        Log
        (
          `Failed to update prior year basis for account <${sheetName}> to <${valuePriorYear}> (it is still <${basisCurrentYear}>).`
        );
      }
    }
    else
    {
      LogVerbose
      (
        `Currently stored prior year basis <${valuePriorYear}> for account <${sheetName}> ` +
        `matches value <${basisCurrentYear}> from previous year.`,
        verbose
      );
    }
  }
};


/**
 * UpdateAllocationSheet()
 *
 * Update history and save important values in the public allocation sheet
 */
function UpdateAllocationSheet(sheetID, scriptTime, verbose, confirmNumbers)
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
    SaveValuesInHistory(sheetID, historySheetName, names, scriptTime, backupRun, updateRun, verbose);
    
    // Preserve RUT performance statistics
    historySheetName = "H: RUT";
    names=
    {
      "RUTPrice" : 0,
      "RUTBufferPuts" : 0,
      "RUTBufferCalls" : 0,
      "RUTProfit" : -5
    };
    SaveValuesInHistory(sheetID, historySheetName, names, scriptTime, backupRun, updateRun, verbose);
    
    // Preserve SPX performance statistics
    historySheetName = "H: SPX";
    names=
    {
      "SPXPrice" : 0,
      "SPXBufferPuts" : 0,
      "SPXBufferCalls" : 0,
      "SPXProfit" : -5
    };
    SaveValuesInHistory(sheetID, historySheetName, names, scriptTime, backupRun, updateRun, verbose);
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
  return "1lB6ndpg16wT14JIOiHtKbs_sdn_ks_qFHJKP5IN6ljg";
};


/**
 * MapNames()
 *
 * Map one set of names to another by adding a suffix and optionally, dropping an existing suffix
 */
function MapNames(sourceNames, suffix, dropSuffix)
{
  const mappedNames= [];
  
  if (dropSuffix != undefined)
  {
    // Map source names to new names by repalcing a current suffix with a new one
    const dropSuffixExpression = new RegExp(dropSuffix + "$", "")
    
    for (const index in sourceNames)
    {
      if (dropSuffixExpression.test(sourceNames[index]))
      {
        // Replace an existing suffix
        mappedNames.push(sourceNames[index].replace(dropSuffixExpression, suffix));
      }
      else
      {
        // Did not match an existing suffix, just add the new one
        mappedNames.push(sourceNames[index] + suffix);
      }
    }
  }
  else
  {
    // Map source names to new names by adding a suffix
    for (const index in sourceNames)
    {
      mappedNames.push(sourceNames[index] + suffix);
    }
  }
  
  return mappedNames;
};