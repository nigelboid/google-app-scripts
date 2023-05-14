/**
 * Main entry point for the continuous check
 *
 * Checks for fresh data and saves it once it is available
 */
function RunGabelliHourly()
{
  // declare constants and local variables
  var id= GetGabelliSheetID();
  var scriptTime= new Date();
  var stampName= "StampHourly";
  var stampDateName= "StampHourlyDate";
  var navDateName= "NAVDate";
  var earliestPossibleUpdate= 1700;
  var latestPossibleUpdate= 2359;
  var verbose= false;
  
  if ((scriptTime.getDay() > 0 && scriptTime.getDay() < 6)
    && ((scriptTime.getHours()*100+scriptTime.getMinutes()) > earliestPossibleUpdate
    && (scriptTime.getHours()*100+scriptTime.getMinutes()) < latestPossibleUpdate))
  {
    SetValueByName(id, stampName, scriptTime, verbose);
    SetValueByName(id, stampDateName, GetValueByName(id, navDateName, verbose), verbose);
    SnapshotOverview(id, scriptTime, verbose);
  }
  
  LogSend(id);
};


/**
 * Main entry point for the daily check
 *
 * Checks for fresh data and saves it once it is available
 */
function RunGabelliDaily(backupRun)
{
  // declare constants and local variables
  var id= GetGabelliSheetID();
  var stampName= "StampDaily";
  var stampDateName= "StampDailyDate";
  var navDateName= "NAVDate";
  var spreadsHistorySheetName= "H: Spreads";
  var spreadsNames= ["Spreads"];
  var spreadsLimits= [-1];
  var now= new Date();
  var verbose= false;
  var updateRun= false;
  
  SetValueByName(id, stampName, now, verbose);
  SetValueByName(id, stampDateName, GetValueByName(id, navDateName, verbose), verbose);
  SaveValuesInHistory(id, spreadsHistorySheetName, spreadsNames, spreadsLimits, now, backupRun, updateRun, verbose);
  
  LogSend(id);
};


/**
 * Main entry point for NAV and price snapshots
 *
 * Saves current NAVs and prices in history per fund
 */
function SnapshotOverview(id, now, verbose)
{
  // declare constants and local variables
  var overview= null;
  var sourceName= "Overview";
  var targetName= null;
  var targetNamePrefix= "H: ";
  var currentDataDate= null;
  var firstDataColumn= 1;
  var storeIterationCount= true;
  var confirmNumbers= true;
  var limit= 0;
  
  // get the overview snapshot (expecting name in first column and numerical data in the rest)
  overview= GetTableByName(id, sourceName, firstDataColumn, confirmNumbers, limit, storeIterationCount, verbose);
  
  // confirm and save current values
  if (overview)
  {
    // we have values, save them
    for (var vIndex= 0; vIndex < overview.length; vIndex++)
    {
      targetName= targetNamePrefix + overview[vIndex].shift();
      
      currentDataDate= overview[vIndex][0];
      if (currentDataDate)
      {
        // obtaimed date string
        currentDataDate= new Date(currentDataDate);
        if (currentDataDate)
        {
          // converted string to date
          if (verbose)
          {
            Logger.log("[SnapshotOverview] Processing data for <%s>.", currentDataDate);
          }
          
          if (!CheckSnapshot(id, targetName, currentDataDate, verbose))
          {
            SaveSnapshot(id, targetName, overview[vIndex], false, verbose);
          }
        }
        else
        {
          Logger.log("[SnapshotOverview] We seem to have data for <%s> without a date stamp -- skipping...", targetName);
        }
      }
      else
      {
        Logger.log("[SnapshotOverview] We seem to have data for <%s> without a date string -- skipping...", targetName);
      }
    }
  }
};


/**
* GetGabelliSheetID()
 *
 * Return a one-dimensional array filled with specified values
 */
function GetGabelliSheetID()
{
  return "1PYzBS3nHLJq6A1DxLKi-ch6JOf62rhx2VtLtNwczb9U";
};