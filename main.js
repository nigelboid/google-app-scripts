/**
 * Wrapper for the main entry point of the script collection
 * should be invoked as the primary run of the day
 *
 * Invokes appropriate Run() methods for individual scripts
 */
function RunPrimary()
{
  var backupRun = false;
  
  RunDaily(backupRun);
};


/**
 * Wrapper for the main entry point of the script collection
 * should be invoked as a backup run
 *
 * Invokes appropriate Run() methods for individual scripts, backing up previous, possibly failed invocations
 */
function RunBackup()
{
  var backupRun = true;
  
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
  // RunLendingClubDaily(backupRun);
};


/**
 * Main entry point for the hourly run script collection
 *
 * Invokes appropriate Run() methods for individual scripts
 */
function RunHourly()
{
  var afterHours = true;
  
  RunQuotes(afterHours);
  RunPersonalHourly();
  RunAuxiliary();
  // RunLendingClubHourly();
};


/**
 * Main entry point for the frequently (several times per hour) run script collection
 *
 * Invokes appropriate Run() methods for individual scripts
 */
function RunFrequently()
{
  var afterHours = false;
  
  if (!RunIndexStranglesCandidates(afterHours))
  {
    // Only update quotes if candidates and boxes skipped (new candidates and boxes will force quotes)
    if (!RunQuotes(afterHours))
    {
      // Only check electricity prices if the rest skipped (try to stay within a narrow execution window)
      RunComEdFrequently();
    }
  }
};