/**
 * Wrapper for the main entry point of the script collection
 * should be invoked as the primary run of the day
 *
 * Saves a bit of data from one spreadsheet in another
 */
function RunPrimary()
{
  var backupRun= false;
  
  RunDaily(backupRun);
};


/**
 * Wrapper for the main entry point of the script collection
 * should be invoked as a backup run
 *
 * Saves a bit of data from one spreadsheet in another
 */
function RunBackup()
{
  var backupRun= true;
  
  RunDaily(backupRun);
};


/**
 * Main entry point for the daily run script collection
 *
 * Invokes appropriate Run() methods for individual scripts
 */
function RunDaily(backupRun)
{
  RunPersonal(backupRun);
  RunLendingClubDaily(backupRun);
  RunBoxTradeCandidates(backupRun)
};


/**
 * Main entry point for the hourly run script collection
 *
 * Invokes appropriate Run() methods for individual scripts
 */
function RunHourly()
{
  var afterHours= true;
  
  RunQuotes(afterHours);
  RunPersonalHourly();
  RunLendingClubHourly();
  RunAuxiliary();
};


/**
 * Main entry point for the frequently (several times per hour) run script collection
 *
 * Invokes appropriate Run() methods for individual scripts
 */
function RunFrequently()
{
  var afterHours= false;
  //var afterHours= true;
  
  if (!RunIndexStranglesCandidates())
  {
    // Only update quotes if candidates skipped (try to stay within a narrow execution window; new candidates will force quotes)
    if (!RunQuotes(afterHours))
    {
      // Only check electricity prices if the rest skipped (try to stay within a narrow execution window)
      RunComEdFrequently();
    }
  }
};


/**
 * Main entry point for testing after hours
 *
 * Invokes appropriate Run() methods for individual scripts
 */
function RunAfterHours()
{
  var afterHours= true;
  
  RunQuotes(afterHours);
};


/**
 * RunCustom()
 *
 * Custom short name wrapper for a longer named function
 *
 */
function RunCustom()
{
  var afterHours= false;

  RunQuotes(afterHours)
};


/**
 * RunTest()
 *
 * Custom short name wrapper for a longer named function
 *
 */
function RunTest()
{
  RunPersonalTest();
};