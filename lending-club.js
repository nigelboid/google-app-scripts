/**
 * Main entry point for the continuous check
 *
 * Checks for fresh data and saves it once it is available
 */
function RunLendingClubHourly()
{
  // Declare constants and local variables
  var id= GetLendingClubSheetID();
  var scriptTime= new Date();
  var interestingHours= [23];
  var verbose= false;
  
  // Poll a value to see if there is a lingering error
  var annualizedReturn= GetValueByName(id, "accountAnnualizedReturn", verbose);
  
  for (const hour of interestingHours)
  {
    // Did we get invoked during interesting times or is there a lingering error from last time?
    if (scriptTime.getHours().toFixed(0) == hour.toFixed(0) || isNaN(annualizedReturn))
    {
      // Yes!
      var summarySheetName= GetValueByName(id, "lcSummarySheet", verbose);
      //var ownedNotesSheetName= GetValueByName(id, "lcNotesSheet", verbose);
      
      if (isNaN(annualizedReturn))
      {
        // Report an error
        Logger.log("[RunLendingClubHourly] Triggering a run outside interesting hours due to a lingering error <%s>", annualizedReturn)
      }
      
      Save(id, summarySheetName, GetSummary(id, scriptTime, verbose), verbose);
      //Save(id, ownedNotesSheetName, GetOwnedNotes(id, scriptTime, verbose), verbose);
      break;
    }
    else
    {
      //Logger.log("[RunLendingClubHourly] These are not interesting times: script hour <%s> versus interesting hour <%s>)",
      //scriptTime.getHours().toFixed(0), hour.toFixed(0));
    }
  }
  
  LogSend(id);
};


/**
 * Main entry point for the daily check
 *
 * Preserve certain data snapshots
 */
function RunLendingClubDaily(backupRun)
{
  // declare constants and local variables
  var id= GetLendingClubSheetID();
  var historySheetName= "History";
  var names= [];
  var savedNames= [];
  var limits= [0, 0, 0, 0, 0, -100000, -100000, -1, -1, -1, -1, -1];
  var now= new Date();
  var verbose= false;
  var verboseChanges= true;
  var updateRun= false;
  
  names= ["accountBasis", "accountValue", "accountInterest", "accountLateFees", "accountInvestorFees", "accountLoss", "accountChargedOff",
              "accountAnnualizedReturn", "accountLCAnnualizedReturn", "accountAdjustedAnnualizedReturn", "accountYields", "accountReturns"];
  SaveValuesInHistory(id, historySheetName, names, limits, now, backupRun, updateRun, verbose);
  
  SaveValue(id, "notesProcessingCurrent", "notesProcessingSaved", verbose);
  SaveValue(id, "notesNonPerformingCurrent", "notesNonPerformingSaved", verbose);
  
  // Synchromnise category counts and carp on changes
  names= ["summaryChargedOff", "summaryPaidOff", "summaryPerforming", "summaryNonPerforming", "accountCash"];
  savedNames= MapNames(names, "Saved");
  if (!Synchronize(id, id, names, savedNames, verbose, verboseChanges))
  {
    Logger.log("[RunLendingClubDaily] Failed to preserve: <%s>", savedNames);
  }
  
  LogSend(id);
};


/**
 * GetSummary()
 *
 * Obtain Lending Club account summary
 */
function GetSummary(id, now, verbose)
{
  // declare constants and local variables
  var timeStamp= ["As of", now];
  var url= GetValueByName(id, "apiSummaryURL", verbose);
  var options= ComposeHeaders(GetValueByName(id, "token", verbose));
  
  var response= UrlFetchApp.fetch(url, options);
  var json= response.getContentText();
  var summary= JSON.parse(json);
  
  var table= [timeStamp];
  for (var line in summary)
  {
    // convert each data line into a two-column row
    table.push([line, summary[line]]);
  }
  
  return table;
};


/**
 * GetOwnedNotes()
 *
 * Obtain Lending Club account summary
 */
function GetOwnedNotes(id, now, verbose)
{
  // declare constants and local variables
  var timeStamp= ["As of", now];
  var myNotes= "myNotes";
  var url= GetValueByName(id, "apiOwnedNotesURL", verbose);
  var options= ComposeHeaders(GetValueByName(id, "token", verbose));
  
  // get a dump of owned notes
  var response= UrlFetchApp.fetch(url, options);
  var json= response.getContentText();
  var notes= JSON.parse(json)[myNotes];
  
  // get the list of columns and transpose them per order specified in the spreadsheet
  var orderedKeysTableName= "notesColumnOrder";
  var orderedKeys= GetTableByNameSimple(id, orderedKeysTableName, verbose);
  var unorderedKeys= Object.keys(notes[0]);
  var keys= [];
  var key= null;
  var position= null;
  for (key of orderedKeys)
  {
    position= unorderedKeys.indexOf(key[0]);
    if (position > -1)
    {
      keys.push(key[0]);
      unorderedKeys.splice(position,1);
    }
    else
    {
      Logger.log("[GetOwnedNotes] Could not find <%s> within data columns returned.", key[0]);
    }
  }
  
  // retain the rest of the columns in their original order
  keys= keys.concat(unorderedKeys);
  
  var table= [keys];
  var values= [];
  var note= null;
  table.unshift(timeStamp.concat(FillArray(table[0].length - timeStamp.length, "")));
  for (note of notes)
  {
    // repackage each note record into an array of values corresponding to the array of keys
    position= 0;
    for (key of keys)
    {
      values[position]= note[key];
      position++;
    }
    table.push(values.slice(0));
  }
  
  return table;
};


/**
 * ComposeHeaders()
 *
 * Compose required headers for our requests
 */
function ComposeHeaders(token)
{
  // declare constants and local variables
  var headers= { 'Authorization': token };
  
  return { 'headers': headers, 'muteHttpExceptions' : true };
};


/**
 * Save()
 *
 * Saves current data in a sheet, replacing prior data
 */
function Save(id, sheetName, data, verbose)
{
  // declare constants and local variables
  var spreadsheet= null;
  var sheet= null;
  var range= null;
   
  if (data)
  {
    // we have viable data to save
    var height= data.length;
    var width= data[0].length;
    if (spreadsheet= SpreadsheetApp.openById(id))
    {
      if (sheet= spreadsheet.getSheetByName(sheetName))
      {
        if (sheet= sheet.clearContents())
        {
          if (range= sheet.getRange(1, 1, height, width))
          {
            range= range.setValues(data)
            if (!range)
            {
              Logger.log("[Save] Could not save data in sheet <%s> for spreadsheet <%s>.", sheetName, spreadsheet.getName());
            }
          }
          else
          {
            Logger.log("[Save] Could not set range to accommodate our data set (<%s> rows by <%s> columns) in sheet <%s> for spreadsheet <%s>.",
                       height.toFixed(0), width.toFixed(0), sheetName, spreadsheet.getName());
          }
        }
        else
        {
          Logger.log("[Save] Could not clear contents of sheet <%s> in spreadsheet <%s>.", sheetName, spreadsheet.getName());
        }
      }
      else
      {
        Logger.log("[Save] Could not activate sheet <%s> for spreadsheet <%s>.", sheetName, spreadsheet.getName());
      }
    }
    else
    {
      Logger.log("[Save] Could not open spreadsheet ID <%s>.", id);
    }
  }
  else
  {
    Logger.log("[Save] Nothing to write to sheet <%s> of spreadsheet ID <%s>...", sheetName, id);
  }
};


/**
* GetLendingClubSheetID()
 *
 * Return a one-dimensional array filled with specified values
 */
function GetLendingClubSheetID()
{
  return "11ha4q56Fnqv76pHv5cD2w-NkM3EL0KnjnPTuiUdtThg";
};