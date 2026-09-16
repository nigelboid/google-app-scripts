/**
 * GetTableByNameSimple()
 *
 * Simplified wrapper for a more flexible GetTableByName()
 */
function GetTableByNameSimple(sheetID, sourceName, verbose)
{
  var firstDataColumn = 0;
  var confirmNumbers = false;
  var limit = null;
  var storeIterationCount = false;
  
  return GetTableByName(sheetID, sourceName, firstDataColumn, confirmNumbers, limit, storeIterationCount, verbose);
};
  
  
/**
 * GetTableByRangeSimple()
 *
 * Simplified wrapper for a more flexible GetTableByRange()
 */
function GetTableByRangeSimple(range, verbose)
{
  var firstDataColumn = 0;
  var confirmNumbers = false;
  var limit = null;
  var storeIterationCount = false;
  
  return GetTableByRange(range, firstDataColumn, confirmNumbers, limit, storeIterationCount, verbose);
};
  
  

/**
 * GetTableByName()
 *
 * Read a table of data into a 2-dimensional array and optionally confirm numeric results
 */
function GetTableByName(sheetID, sourceName, firstDataColumn, confirmNumbers, limit, storeIterationCount, verbose)
{
  var spreadsheet = null;
  var range = null;
  var table = null;
  
  if (spreadsheet = SpreadsheetApp.openById(sheetID))
  {
    if (range = spreadsheet.getRangeByName(sourceName))
    {
      table = GetTableByRange(range, firstDataColumn, confirmNumbers, limit, storeIterationCount, verbose);
    }
    else
    {
      LogVerbose(`Could not get range named <${sourceName}> in spreadsheet <${spreadsheet.getName()}>.`, verbose);
    }
  }
  else
  {
    LogVerbose(`Could not open spreadsheet ID <${sheetID}>.`, verbose);
  }
  
  return table;
};
  
  
/**
 * GetTableByRange()
 *
 * Read a range of data into a 2-dimensional array and optionally confirm numeric results
 */
function GetTableByRange(range, firstDataColumn, confirmNumbers, limit, storeIterationCount, verbose)
{
  var data = null;
  var table = [];
  var good = true;
  var maxIterations = 10;
  var sleepInterval = 5000;
  var iterationErrors = null;
  
  for (var iteration = 1; iteration <= maxIterations; iteration++)
  {
    iterationErrors = [];
    data = range.getValues();
    if (data)
    {
      for (var vIndex = 0; vIndex < data.length; vIndex++)
      {
        if (confirmNumbers)
        {
          for (var hIndex = firstDataColumn; hIndex < data[vIndex].length; hIndex++)
          {
            // check each value obtained to make sure it is a number and above the sanity check limit
            if ((data[vIndex][hIndex] == null) || isNaN(data[vIndex][hIndex]) || (data[vIndex][hIndex] < limit))
            {
              iterationErrors.push
              (
                `Could not get a viable value (<${data[vIndex][hIndex]}> v. limit of <${limit}>) from location ` +
                `<${hIndex.toFixed(0)}, ${vIndex.toFixed(0)}>.`
              );
              
              data[vIndex][hIndex] = iteration.toFixed(0);
              good = false;
            }
          }
        }
        
        if (good)
        {
          // All values checked out against the limit -- save current data row and append current iteration count (if asked)
          table[vIndex] = data[vIndex];
          if (storeIterationCount)
          {
            table[vIndex].push(iteration.toFixed(0));
          }
        }
      }
      
      if (good)
      {
        // All values checked out against the limit -- we're done here
        break;
      }
      else
      {
        // Reset the flag -- we'll try again
        good = true;
      }
    }
    else
    {
      LogVerbose("Could not read data from range.", verbose);
      data = iteration;
    }
    
    Utilities.sleep(sleepInterval);
  }
  
  if (iterationErrors.length > 0)
  {
    // Encountered errors while reading data -- report them
    while(iterationErrors.length > 0)
    {
      // Report all the accumulated errors
      Log(iterationErrors.shift());
    }
    Log(`Reached <${iteration.toFixed(0)}> iterations but still could not get all data from range.`);

    table = null;
  }
  
  return table;
};


/**
 * GetValueByName()
 *
 * Obtain a value from a labeled one-cell range
 */
function GetValueByName(sheetID, sourceName, verbose, confirmNumbers, limit)
{
  var value = null;
  var firstDataColumn =  0;
  var storeIterationCount = false;
  
  if (confirmNumbers == undefined)
  {
    confirmNumbers = false;
    limit = 0;
    LogVerbose(`Defaulting: Not confirming numbers with limit set to <${limit}>.`, verbose);
  }
  else
  {
    if (confirmNumbers)
    {
      // make sure limit is defined if we are to confirm numbers
      if (limit == undefined)
      {
        limit = 0;
        LogVerbose(`Defaulting: Limit set to <${limit}>.`, verbose);
      }
    }
  }
  
  value = GetTableByName(sheetID, sourceName, firstDataColumn, confirmNumbers, limit, storeIterationCount, verbose);
  if (value)
  {
    // We seem to have something!
    if (value.length > 0)
    {
      // We seem to have at least one dimension!
      if (value[0].length > 0)
      {
        // We seem to have a proper table -- assign the top-left value
        value = value[0][0];
      }
      else
      {
        // Not a proper table!
        LogVerbose(`Range named <${sourceName}> is not a table.`, verbose);
        value = null;
      }
    }
    else
    {
      // Not even a proper array!
      LogVerbose(`Range named <${sourceName}> is not even an array.`, verbose);
      value = null;
    }
  }
  else
  {
    // We got nothing!
    LogVerbose(`Range named <${sourceName}> did not result in a viable value.`, verbose);
    value = null;
  }

  return value;
};


/**
 * SetTableByName()
 *
 * Write a 2-dimensional array of data into a named spreadsheet table
 */
function SetTableByName(sheetID, destinationName, table, verbose)
{
  var spreadsheet = null;
  var range = null;
  var height = null;
  var width = null;
  var success = true;
  
  if (spreadsheet = SpreadsheetApp.openById(sheetID))
  {
    if (range = spreadsheet.getRangeByName(destinationName))
    {
      // perform capacity checks before writing
      height = range.getHeight();
      if (height < table.length)
      {
        LogVerbose
        (
          `Could not write out range named <${destinationName}> in spreadsheet <${spreadsheet.getName()}> ` +
          `since the destination is shorter than the data we have by <${table.length - height}> rows.`,
          verbose
        );
                  
        success= false;
      }
      else
      {
        // looks like we have sufficient height available, now check width
        width = range.getWidth();
        if (width < table[0].length)
        {
          LogVerbose
          (
            `Could not write out range named <${destinationName}> in spreadsheet <${spreadsheet.getName()}> ` +
            `since the destination is narrower than the data we have by <${table[0].length - width}> columns.`,
            verbose
          );

          success= false;
        }
        else
        {
          // looks we also have sufficient width available, pad source data (if necessary)
          for (var vIndex = 0; vIndex < height; vIndex++)
          {
            if ((vIndex + 1) > table.length)
            {
              // add blank rows to the source table
              table.push(FillArray(width, ""));
            }
            else if (width > table[vIndex].length)
            {
              // add blank columns to existing rows
              table[vIndex] = table[vIndex].concat(FillArray(width - table[vIndex].length, ""));
            }
          }
          
          // write out the values
          range = range.setValues(table);
          if (!range)
          {
            LogVerbose(`Could not write out range named <${destinationName}> in spreadsheet <${spreadsheet.getName()}>.`, verbose);
                    
            success = false;
          }
        }
      }
    }
    else
    {
      LogVerbose(`Could not get range named <${destinationName}> in spreadsheet <${spreadsheet.getName()}>.`, verbose);
                      
      success = false;
    }
  }
  else
  {
    LogVerbose(`Could not open spreadsheet ID <${sheetID}>.`, verbose);
    
    success = false;
  }
  
  return success;
};


/**
 * SetValueByName()
 *
 * Set a value from to labeled one-cell range
 */
function SetValueByName(sheetID, destinationName, value, verbose)
{
  return SetTableByName(sheetID, destinationName, [[value]], verbose)
};


/**
 * GetAnnualSheetIDs()
 *
 * Looks up IDs of all known annual sheets
 */
function GetAnnualSheetIDs(sheetID, verbose)
{
  var sourceName = "ExternalLookups";
  var idsByYear = [];
  var sheetIDs = {};
  
  idsByYear = GetTableByNameSimple(sheetID, sourceName, verbose);
  
  if (idsByYear)
  {
    // we have viable IDs
    for (var vIndex = 0; vIndex < idsByYear.length; vIndex++)
    {
      sheetIDs[idsByYear[vIndex][0]] = idsByYear[vIndex][1];
    }
  }
  else
  {
    LogVerbose(`Could not obtain a list of annual sheet IDs from table <${sourceName}> of spreadsheet ID <${sheetID}>.`, verbose);
  }
  
  return sheetIDs;
};


/**
 * SaveValue()
 *
 * Save current values in a mirror table
 */
function SaveValue(sheetID, sourceName, destinationName, verbose, confirmNumbers, limit)
{
  var sourceValues = [];
  var destinationValues = [];
  var firstDataColumn = 0;
  var storeIterationCount = false;
  var changed = false;
  
  // set defaults unless supplied
  if (confirmNumbers == undefined)
  {
    confirmNumbers = false;
    limit = 0;

    LogVerbose(`Not confirming numbers...`, verbose);
  }
  else
  {
    if (confirmNumbers)
    {
      // make sure limit is defined if we are to confirm numbers
      if (limit == undefined)
      {
        limit = 0;

        LogVerbose(`Confirming numbers with a limit set to <${limit}>...`, verbose);
      }
    }
  }
  
  // Read all the source and destination values, compare, and update
  if (sourceValues = GetTableByName(sheetID, sourceName, firstDataColumn, confirmNumbers, limit, storeIterationCount, verbose))
  {
    // we have source values, proceed to destination values
    if (destinationValues = GetTableByName(sheetID, destinationName, firstDataColumn, confirmNumbers, limit, storeIterationCount, verbose))
    {
      // compare values and update them
      if (sourceValues.length == destinationValues.length)
      {
        for (var vIndex= 0; vIndex < destinationValues.length; vIndex++)
        {
          if (sourceValues[vIndex].length == destinationValues[vIndex].length)
          {
            for (var hIndex= 0; hIndex < destinationValues[vIndex].length; hIndex++)
            {
              if (sourceValues[vIndex][hIndex] != destinationValues[vIndex][hIndex])
              {
                LogVerbose
                (
                  `Value at location <${hIndex.toFixed(0)}, ${vIndex.toFixed(0)}> ` +
                  `has changed to <${sourceValues[vIndex][hIndex]}> in table <${sourceName}> ` +
                  `from <${destinationValues[vIndex][hIndex]}> in table <${destinationName}> ` +
                  `of spreadsheet ID <${sheetID}>.`,
                  verbose
                );
                    
                destinationValues[vIndex][hIndex] = sourceValues[vIndex][hIndex];
                changed = true;
              }
              else
              {
                LogVerbose
                (
                  `Value <${destinationValues[vIndex][hIndex]}> (${sourceValues[vIndex][hIndex]}) ` +
                  `at location <${hIndex.toFixed(0)}, ${vIndex.toFixed(0)}> ` +
                  `has not changed between named tables <${sourceName}> and <${destinationName}>  of spreadsheet ID <${sheetID}>.`,
                  verbose
                );
              }
            }
          }
          else
          {
            Log
            (
              `Source values range <${sourceValues.length}, ${sourceValues[vIndex].length}> of source <${sourceName}> ` +
              `does not match destination range <${destinationValues.length}, ${destinationValues[vIndex].length}> ` +
              `of destination <${destinationName}> in spreadsheet ID <${sheetID}>.`
            );
          }
        }
      }
      else
      {
        Log
        (
          `Source values height <${sourceValues.length}> of source <${sourceName}> ` +
          `does not match destination height <${destinationValues.length}> of destination <${destinationName}> ` +
          `in spreadsheet ID <${sheetID}>.`
        );
      }
      
      if (changed)
      {
        // write out the values
        if (!SetTableByName(sheetID, destinationName, destinationValues, verbose))
        {
          // something went wrong!
          Log(`Could not write out range named <${destinationName}> in spreadsheet ID <${sheetID}>.`);
          
          changed = false;
        }
      }
    }
    else
    {
      Log(`Could not get range named <${destinationName}> in spreadsheet ID <${sheetID}>.`);
    }
  }
  else
  {
    Log(`Could not get range named <${sourceName}> in spreadsheet ID <${sheetID}>.`);
  }
  
  return changed;
};


/**
 * GetLastSnapshotStamp()
 *
 * Obtain the identifier stamp for the last snapshot entry
 */
function GetLastSnapshotStamp(sheetID, sheetName, verbose)
{
  var spreadsheet = null;
  var sheet = null;
  var range = null;
  var height = null;
  var value = null;
  
  if (spreadsheet = SpreadsheetApp.openById(sheetID))
  {
    if (sheet = spreadsheet.getSheetByName(sheetName))
    {
      if (height = sheet.getLastRow())
      {
        if (range = sheet.getRange(height, 1))
        {
          value = range.getValue();
        }
        else
        {
          LogVerbose
          (
            `Could not set range to the first cell of the last row <${height.toFixed(0)}> ` +
            `in sheet <${sheetName()}> for spreadsheet <${spreadsheet.getName()}>.`,
            verbose
          );
        }
      }
      else
      {
        LogVerbose(`Could not learn the last row in sheet <${sheetName()}> for spreadsheet <${spreadsheet.getName()}>.`, verbose);
      }
    }
    else
    {
      LogVerbose(`Could not activate sheet <${sheetName()}> for spreadsheet <${spreadsheet.getName()}>.`, verbose);
    }
  }
  else
  {
    LogVerbose(`Could not open spreadsheet ID <${sheetID}>.`, verbose);
  }
  
  return value;
};


/**
 * SelectCell()
 *
 * Obtain a range specification for a given cell of the named sheet
 */
function SelectCell(sheetID, sheetName, cellCoordinates, verbose)
{
  var spreadsheet = null;
  var sheet = null;
  var range = null;
  
  if (spreadsheet = SpreadsheetApp.openById(sheetID))
  {
    if (sheet = spreadsheet.getSheetByName(sheetName))
    {
      range = sheet.getRange(cellCoordinates);
    }
    else
    {
      Log(`Could not activate sheet <${sheetName()}> for spreadsheet <${spreadsheet.getName()}>.`);
    }
  }
  else
  {
    Log(`Could not open spreadsheet ID <${sheetID}>.`);
  }

  return range;
};


/**
 * GetCellValue()
 *
 * Obtain the value of the specified cell
 */
function GetCellValue(sheetID, sheetName, cellCoordinates, verbose)
{
  var range = SelectCell(sheetID, sheetName, cellCoordinates, verbose);
  var value = null;
  
  if (range)
  {
    value= range.getValue();
  }
  else
  {
    LogVerbose(`Could not access specified range <${cellCoordinates}> in sheet <${sheetName}> for spreadsheet ID <${sheetID}>.`, verbose);
  }
  
  return value;
};


/**
 * SetCellValue()
 *
 * Set the value of the specified cell
 */
function SetCellValue(sheetID, sheetName, cellCoordinates, value, verbose)
{
  var range = SelectCell(sheetID, sheetName, cellCoordinates, verbose);
  var success = false;
  
  if (range)
  {
    success = range.setValue(value);
  }
  else
  {
    LogVerbose(`Could not access specified range <${cellCoordinates}> in sheet <${sheetName}> for spreadsheet ID <${sheetID}>.`, verbose);
  }
  
  return success;
};


/**
 * CheckSnapshot()
 *
 * Check data in the destination spreadsheet
 */
function CheckSnapshot(sheetID, sheetName, newDataDate, verbose)
{
  var lastDataDate = new Date(GetLastSnapshotStamp(sheetID, sheetName, verbose));
  
  if (lastDataDate && (lastDataDate.getFullYear() == newDataDate.getFullYear()) &&
     (lastDataDate.getMonth() == newDataDate.getMonth()) && (lastDataDate.getDate() == newDataDate.getDate()))
  {
    // we have already recorded data for today
    return true;
  }
  else if (lastDataDate > newDataDate)
  {
    Log(`We seem to have stale data from the past (last date <${lastDataDate}> is later than new date <${newDataDate}>, skipping...`);

    return true;
  }
  else
  {
    // we don't have the latest data
    return false;
  }
};


/**
 * CompileSnapshot()
 *
 * Compile our snapshot from various cells
 */
function CompileSnapshot(sheetID, names, dateTimeNow, verbose)
{
  const firstDataColumn = 0;
  const storeIterationCount = true;
  var confirmNumbers = false;
  var good = true;
  var iterations = 0;
  var snapshot = [dateTimeNow];
  var table = null;
  
  for (const name in names)
  {
    // check whether we should confirm numeric values for this data point
    if (Number.isNaN(names[name]) || typeof names[name] != "number")
    {
      confirmNumbers = false;
    }
    else
    {
      confirmNumbers = true;
    }

    // read each table or cell of interest and accumulate in an array
    table = GetTableByName(sheetID, name, firstDataColumn, confirmNumbers, names[name], storeIterationCount, verbose);
    if (table)
    {
      // Got viable data -- now transpose the table
      for (const row in table)
      {
        // Grab the first column value from every row returned
        snapshot.push(table[row][0]);
        if (iterations < table[row][table[row].length - 1])
        {
          // Store the highest iteration count
         iterations = table[row][table[row].length - 1];
        }
      }
    }
    else
    {
      // Failed to obtain viable data
      good = false;

      const spreadsheet = SpreadsheetApp.openById(sheetID);
      if (spreadsheet)
      {
        const sheetName = spreadsheet.getName();

        if (sheetName)
        {
          Log(`Failed to read data for <${name}> from sheet <${sheetName}>, skipping the entire snapshot...`);
        }
        else
        {
          Log(`Failed to read data for <${name}> from sheet ID <${sheetID}>, skipping the entire snapshot...`);
          Log(`Also, failed to obtain the name of the sheet.`);
        }
      }
      else
      {
        Log(`Failed to read data for <${name}> from sheet ID <${sheetID}>, skipping the entire snapshot...`);
        Log(`Also, failed to open the sheet by ID.`);
      }

      break;
    }
  }
  
  if (good)
  {
    // dress and return viable data
    snapshot[0] = new Date();
    snapshot.push(iterations);
  }
  else
  {
    // No viable data to return
    snapshot = null;
  }

  return snapshot;
};


/**
 * SaveSnapshot()
 *
 * Save values snapshot in a history table
 */
function SaveSnapshot(sheetID, sheetName, values, updateRun, verbose)
{
  var spreadsheet = null;
  var sheet = null;
  var range = null;
  var lastRow = null;
  var success = false;
  
  if (values)
  {
    // we have viable data to save
    if (!Array.isArray(values[0]))
    {
      // we seem to have a one-dimensional array -- convert it
      values = [values];
    }
    
    // now access the spreadsheet and save
    if (spreadsheet = SpreadsheetApp.openById(sheetID))
    {
      if (sheet = spreadsheet.getSheetByName(sheetName))
      {
        if (lastRow = sheet.getLastRow())
        {
          if (!updateRun)
          {
            // this is not an update run -- append a row
            lastRow++;
          }
          
          if (range = sheet.getRange(lastRow, 1, values.length, values[0].length))
          {
            if (range = range.setValues(values))
            {
              if (PropagateFormulas(sheet, lastRow, values[0].length, verbose))
              {
                success = true;
              }
              else
              {
                LogVerbose(`Could not propagate formulas in sheet <${sheetName}> for spreadsheet <${spreadsheet.getName()}>.`, verbose);
              }
            }
            else
            {
              LogVerbose(`Could not append values in sheet <${sheetName}> for spreadsheet <${spreadsheet.getName()}>.`, verbose);
            }
          }
          else
          {
            LogVerbose
            (
              `Could not set range to append beyond the last row <${height.toFixed(0)}> ` +
              `in sheet <${sheetName}> for spreadsheet <${spreadsheet.getName()}>.`,
              verbose
            );
          }
        }
        else
        {
          LogVerbose(`Could not learn the last row in sheet <${sheetName}> for spreadsheet <${spreadsheet.getName()}>.`, verbose);
        }
      }
      else
      {
        LogVerbose(`Could not activate sheet <${sheetName}> for spreadsheet <${spreadsheet.getName()}>.`, verbose);
      }
    }
    else
    {
      LogVerbose(`Could not open spreadsheet ID <${sheetID}>.`, verbose);
    }
  }
  else
  {
    LogVerbose(`Nothing to write to sheet <${sheetName}> of spreadsheet ID <${sheetID}>.`, verbose);
  }
  
  return success;
};


/**
 * PropagateFormulas()
 *
 * Propagate formulas from the row above
 */
function PropagateFormulas(sheet, row, column, verbose)
{
  var width = null;
  var formulas = null;
  // var range = null;
  var success = false;
  
  if (width = sheet.getLastColumn())
  {
    if (width > column)
    {
      // looks like we have spare columns to check
      if (formulas = sheet.getRange(row-1, column+1, 1, width-column).getFormulas())
      {
        if (success = sheet.getRange(row, column+1, 1, width-column).setFormulas(formulas))
        {
          LogVerbose
          (
            `Updated formulas in columns <${(column+1).toFixed(0)}> through <${width.toFixed(0)}> ` +
            `of row <${row.toFixed(0)}> in sheet <${sheet.getName()}>.`,
            verbose
          );
        }
        else
        {
          LogVerbose
          (
            `Could not set formulas in columns <${(column+1).toFixed(0)}> through <${width.toFixed(0)}> ` +
            `of row <${row.toFixed(0)}> in sheet <${sheet.getName()}>.`,
            verbose
          );
        }
      }
      else
      {
        LogVerbose
        (
          `Could not read formulas from columns <${(column+1).toFixed(0)}> through <${width.toFixed(0)}> ` +
          `of row <${(row-1).toFixed(0)}> in sheet <${sheet.getName()}>.`,
          verbose
        );
      }
    }
    else
    {
      LogVerbose(`No columns to propagate in sheet <${sheet.getName()}>.`, verbose);
      
      success = true;
    }
  }
  else
  {
    LogVerbose(`Could not obtain width of sheet <${sheet.getName()}>.`, verbose);
  }
  
  return success;
};


/**
 * UpdateSnapshotCell()
 *
 * Update a specific value in a history table
 */
function UpdateSnapshotCell(sheetID, sheetName, column, value, onlyIfBlank, verbose)
{
  var spreadsheet = null;
  var sheet = null;
  var range = null;
  var height = null;
  var success = false;
  
  if (value != null)
  {
    // now access the spreadsheet and save
    if (spreadsheet = SpreadsheetApp.openById(sheetID))
    {
      if (sheet = spreadsheet.getSheetByName(sheetName))
      {
        if (height = sheet.getLastRow())
        {
          if (range = sheet.getRange(height, column, 1, 1))
          {
            if (!onlyIfBlank || range.isBlank())
            {
              if (range = range.setValue([[value]]))
              {
                LogVerbose
                (
                  `Updated cell <${column.toFixed(0)}> of the last row <${height.toFixed(0)}> ` +
                  `in sheet <${sheetName}> for spreadsheet <${spreadsheet.getName()}> with <${value}>.`,
                  verbose
                );
                    
                success = true;
              }
              else
              {
                LogVerbose
                (
                  `Could not updated cell <${column.toFixed(0)}> of the last row <${height.toFixed(0)}> ` +
                  `in sheet <${sheetName}> for spreadsheet <${spreadsheet.getName()}>.`,
                  verbose
                );
              }
            }
            else
            {
              LogVerbose
              (
                `Could not updated cell <${column.toFixed(0)}> of the last row <${height.toFixed(0)}> ` +
                `in sheet <${sheetName}> for spreadsheet <${spreadsheet.getName()}> ` +
                `since that would clobber an existing value <${range.getValue()}>.`,
                verbose
              );
            }
          }
          else
          {
            LogVerbose
            (
              `Could not set range to update the last row <${height.toFixed(0)}> ` +
              `in sheet <${sheetName}> for spreadsheet <${spreadsheet.getName()}>.`,
              verbose
            );
          }
        }
        else
        {
          LogVerbose(`Could not learn the last row in sheet <${sheetName}> for spreadsheet <${spreadsheet.getName()}>.`, verbose);
        }
      }
      else
      {
        LogVerbose(`Could not activate sheet <${sheetName}> for spreadsheet <${spreadsheet.getName()}>.`, verbose);
      }
    }
    else
    {
      LogVerbose(`Could not open spreadsheet ID <${sheetID}>.`, verbose);
    }
  }
  else
  {
    LogVerbose(`Nothing to update in column <${column}> of sheet <${sheetName}> for spreadsheet <${spreadsheet.getName()}>.`, verbose);
  }
  
  return success;
};


/**
 * SaveValuesInHistory()
 *
 * Saves current values in a history table
 */
function SaveValuesInHistory(sheetID, sheetName, sourceNames, now, backupRun, updateRun, verbose)
{
  if (CheckSnapshot(sheetID, sheetName, now, verbose))
  {
    // History exists for today
    if (updateRun)
    {
      SaveSnapshot(sheetID, sheetName, CompileSnapshot(sheetID, sourceNames, now, verbose), updateRun, verbose);
    }
    else if (!backupRun)
    {
      Log(`Redundant primary run at <${now}> for sheet <${sheetName}> in spreadsheet ID ${sheetID}>.`);
    }
  }
  else
  {
    // History does not exist for today
    if (backupRun)
    {
      Log(`Primary run seems to have failed for sheet <${sheetName}> in spreadsheet ID <${sheetID}>...`);
    }
    
    SaveSnapshot(sheetID, sheetName, CompileSnapshot(sheetID, sourceNames, now, verbose), false, verbose);
  }
};


/**
 * RemoveDuplicateSnapshot()
 *
 * Remove a recent duplicate entry from the history table
 */
function RemoveDuplicateSnapshot(sheetID, sheetName, verbose)
{
  var spreadsheet = null;
  var sheet = null;
  var range = null;
  var height = null;
  var width = null;
  var rowData = null;
  var ultimateStamp = null;
  var penultimateStamp = null;
  var priorStamp = null;
  
  if (spreadsheet = SpreadsheetApp.openById(sheetID))
  {
    if (sheet = spreadsheet.getSheetByName(sheetName))
    {
      if (height = sheet.getLastRow())
      {
        // Learn the latest time stamp
        if (range = sheet.getRange(height, 1))
        {
          ultimateStamp = new Date(range.getValue());
        }
        else
        {
          Log
          (
            `Could not set range to the first cell of the last row <${height.toFixed(0)}> ` +
            `in sheet <${sheetName}> for spreadsheet <${spreadsheet.getName()}>.`
          );
        }
        
        
        // Learn the prior time stamp (seemingly real and accurate data two rows above latest)
        if (range = sheet.getRange(height-2, 1))
        {
          priorStamp = new Date(range.getValue());
        }
        else
        {
          Log
          (
            `Could not set range to the first cell of the prior row <${(height-2).toFixed(0)}> ` +
            `in sheet <${sheetName}> for spreadsheet <${spreadsheet.getName()}>.`
          );
        }
        
        // Learn the time stamp just before the latest
        if (width = sheet.getLastColumn())
        {
          if (range = sheet.getRange(height-1, 1, 1, width))
          {
            rowData = range.getValues();
            penultimateStamp = new Date(rowData[0][0]);
          }
          else
          {
            Log
            (
              `Could not set range to the first cell of the next to the last row <${(height-1).toFixed(0)}> ` +
              `in sheet <${sheetName}> for spreadsheet <${spreadsheet.getName()}>.`
            );
          }
        }
        else
        {
          Log(`Could not learn the last column in sheet <${sheetName}> for spreadsheet <${spreadsheet.getName()}>.`);
        }
        
        // Remove the next to the last row if its time stamp matches that of the last row or the prior row preceding it
        if (ultimateStamp.getTime() === penultimateStamp.getTime() || priorStamp.getTime() === penultimateStamp.getTime())
        {
          try
          {
            sheet = sheet.deleteRow(height-1)
          }
          catch (error)
          {
            Log("Failed to remove duplicate row:\n".concat(error));
          }
          
          if (sheet)
          {
            return [[priorStamp, "Kept"], rowData[0], [ultimateStamp, "Kept"]];
          }
          else
          {
            Log
            (
              `Failed to remove the penultimate row for time stamp <${penultimateStamp}> ` +
              `in sheet <${sheetName}> for spreadsheet <${spreadsheet.getName()}>.`
            );
          }
        }
        else
        {
          LogVerbose
          (
            `No need to remove history rows as time stamps (<${ultimateStamp}> and <${penultimateStamp}>) ` +
            `in sheet <${sheetName}> for spreadsheet <${spreadsheet.getName()}>.`,
            verbose
          );
        }
      }
      else
      {
        Log(`Could not learn the last row in sheet <${sheetName}> for spreadsheet <${spreadsheet.getName()}>.`);
      }
    }
    else
    {
      Log(`Could not activate sheet <${sheetName}> for spreadsheet <${spreadsheet.getName()}>.`);
    }
  }
  else
  {
    Log(`Could not open spreadsheet ID <${sheetID}>.`);
  }
  
  return false;
};


/**
 * TrimHistory()
 *
 * Remove earliest entries from the history table
 */
function TrimHistory(sheetID, sheetName, maxRows, verbose)
{
  var spreadsheet = null;
  var sheet = null;
  var height = 0;
  
  if (spreadsheet = SpreadsheetApp.openById(sheetID))
  {
    if (sheet = spreadsheet.getSheetByName(sheetName))
    {
      height = sheet.getLastRow();
      
      // Accommodate the header row!
      if (height > (maxRows + 1))
      {
        try
        {
          sheet.deleteRows(2, height - (maxRows + 1));
        }
        catch (error)
        {
          Log("Failed to trim history rows:\n".concat(error));
        }
      }
    }
    else
    {
      Log(`Could not activate sheet <${sheetName}> for spreadsheet <${spreadsheet.getName()}>.`);
    }
  }
  else
  {
    Log(`Could not open spreadsheet ID <${sheetID}>.`);
  }
};


/**
 * Synchronize()
 *
 * Preserve value from one range in another range (designed to work on first element of each range)
 */
function Synchronize(sourceID, destinationID, sourceNames, destinationNames, verbose, verboseChanges, numbersOnly)
{
  var spreadsheet = null;
  var range = null;
  var value = null;
  var format = null;
  var sourceValues = [];
  var success = true;
  
  if (destinationNames == undefined)
  {
    // Omitted destination names indicate ientical names in the target sheet
    destinationNames = sourceNames;
  }
  
  if (verboseChanges == undefined)
  {
    // Initialize omitted optional verbose flag to match overall verbose flag
    verboseChanges = verbose;
  }
  
  if (numbersOnly == undefined)
  {
    // Omitted numbers only flag implies otherwise
    numbersOnly = false;
  }
  
  if (spreadsheet = SpreadsheetApp.openById(sourceID))
  {
    for (const index in sourceNames)
    {
      // read all the source values
      if (range = spreadsheet.getRangeByName(sourceNames[index]))
      {
        value = range.getValue();
        format = range.getNumberFormat();
        if (Number.isNaN(value) || typeof value != "number")
        {
          if (verbose || numbersOnly)
          {
            Log
            (
              `Failed to obtain numeric value <${value}> with format <${format}> from range named <${sourceNames[index]}> ` +
              `in spreadsheet <${spreadsheet.getName()}>.`
            );
          }
          
          if (numbersOnly)
          {
            // Fail gently
            sourceValues.push(null);
            success = false;
          }
          else
          {
            sourceValues.push(value);
          }
        }
        else
        {
          if (format.indexOf("$") > -1)
          {
            // Round all currency amounts to cents
            sourceValues.push(value.toFixed(2));
          }
          else
          {
            sourceValues.push(value);
          }
        }
      }
      else
      {
        LogVerbose(`Could not get range named <${sourceNames[index]}> in spreadsheet <${spreadsheet.getName()}>.`, verbose);
        success = false;
      }
    }
  }
  else
  {
    LogVerbose(`Could not open spreadsheet ID <${sourceID}>.`, verbose);
    success = false;
  }
  
  if (sourceValues.length == destinationNames.length)
  {
    if (spreadsheet = SpreadsheetApp.openById(destinationID))
    {
      for (const index in destinationNames)
      {
        // Compare and write values
        if (range = spreadsheet.getRangeByName(destinationNames[index]))
        {
          value = range.getValue();
          
          if ((sourceValues[index] != null) && (value != sourceValues[index]))
          {
            // Looks like the value has changed -- update it
            range.setValue(sourceValues[index]);
            LogVerbose
            (
              `Value for range <${destinationNames[index]}> in sheet <${spreadsheet.getName()}> updated to <${sourceValues[index]}>, ` +
              `it was <${value}>.`,
              verboseChanges
            );
          }
          else
          {
            LogVerbose(`Value for range <${destinationNames[index]}> has not changed <${sourceValues[index]}>.`, verbose);
          }
        }
        else
        {
          LogVerbose(`Could not get range named <${destinationNames[index]}> in spreadsheet <${spreadsheet.getName()}>.`, verbose);
          success = false;
        }
      }
    }
    else
    {
      LogVerbose(`Could not open spreadsheet ID <${destinationID}>.`, verbose);
      success = false;
    }
  }
  else
  {
    LogVerbose
    (
      `Source values range <${sourceValues.length.toFixed(0)}> ` +
      `does not match destination range <${destinationNames.length.toFixed(0)}>.`,
      verbose
    );
    success = false;
  }
  
  return success;
};


/**
 * GetParameters()
 *
 * Read specified table and return an associative array comprised of key-value pairs from the first two columns
 */
function GetParameters(sheetID, sourceName, verbose)
{
  var parameters = {"sheetID": sheetID, "verbose": verbose};
  var firstDataColumn = 1;
  var confirmNumbers = false;
  var limit = 0;
  var storeIterationCount = false;
  var table = GetTableByName(sheetID, sourceName, firstDataColumn, confirmNumbers, limit, storeIterationCount, verbose);
  
  if (table)
  {
    // We seem to have something!
    if (table.length > 0)
    {
      // We seem to have at least one dimension!
      if (table[0].length >= 2)
      {
        // We seem to have at least two columns
        for (var row= 0; row < table.length; row++)
        {
          // Check each row for a viable key-value pair and preserve them in our associative array
          if (table[row][0] != null && table[row][1] != null)
          {
            // We seem to have a viable key-value pair
            parameters[table[row][0]] = table[row][1];
            parameters[table[row][0].toLowerCase()]= table[row][1];
          }
        }
      }
      else
      {
        // Not a proper table!
        LogVerbose(`Range named <${sourceName}> is not a table.`, verbose);
      }
    }
    else
    {
      // Not even a proper array!
      LogVerbose(`Range named <${sourceName}> is not even an array.`, verbose);
    }
  }
  else
  {
    // We got nothing!
    LogVerbose(`Range named <${sourceName}> did not result in a viable value.`, verbose);
  }
  
  return parameters;
};


/**
 * GetMainSheetID()
 *
 * Return the ID of the main sheet
 */
function GetMainSheetID()
{
  //return SpreadsheetApp.getActiveSpreadsheet().getId();
  
  return "1wCsEJveLzU1s4tKi319VJLQGC8f2cDQVxa036B4aWvo";
};


/**
 * FillArray()
 *
 * Return a one-dimensional array filled with specified values
 */
function FillArray(size, value)
{
  var fill = [];
  var counter = 0;
  
  while (counter < size)
  {
    fill[counter++] = value;
  }
  
  return fill;
};


/**
 * NumberToString()
 *
 * Return a formatted number as a string
 */
function NumberToString(number, width, pad)
{
  var formattedNumber = "" + number;
  
  while (formattedNumber.length < width)
  {
    formattedNumber= pad + formattedNumber;
  }
  
  return formattedNumber;
};


/**
 * SpecifyColumnA1Notation()
 *
 * Return a letter code for a specified spreadsheet column
 */
function SpecifyColumnA1Notation(column, verbose)
{
  var specification = null;
  const firstColumn = "A";
      
  if (column > ("Z".charCodeAt(0) - firstColumn.charCodeAt(0) + 1))
  {
    LogVerbose(`No support (yet) for managing a column <${column}> beyond position 'Z.'`, verbose);
  }
  else
  {
    specification= String.fromCharCode(firstColumn.charCodeAt(0) + column);
  }
  
  return specification;
};


/**
 * DateToLocaleString()
 *
 * Return a formatted date as a short string (m/dd/yyyy hh:mm:ss)
 */
function DateToLocaleString(date, separator)
{
  var dateOptions = { day: '2-digit', month: '2-digit', year: 'numeric' };
  //var timeOptions = { hour12: false, hourCycle: 'h23', hour: '2-digit', minute:'2-digit', second: '2-digit'};
  var timeOptions = { hourCycle: 'h23', hour: '2-digit', minute:'2-digit', second: '2-digit'};
  
  if (date == undefined)
  {
    date = new Date();
  }
  else
  {
    date = new Date(date);
  }
  
  if (separator == undefined)
  {
    separator = " ";
  }
  
  //return date.toLocaleString('en-US', {hour12: false, hourCycle: 'h23'});
  //return date.toLocaleString('en-US', {hourCycle: 'h23'});
  return date.toLocaleDateString('en-US', dateOptions) + separator + date.toLocaleTimeString('en-US', timeOptions)
};


/**
 * UpdateTime()
 *
 * Record update time
 */
function UpdateTime(sheetID, timeStampName, verbose)
{
  SetValueByName(sheetID, timeStampName, DateToLocaleString(), verbose);
};


/**
 * isDate()
 *
 * Validate a value as a date
 */
function isDate(value)
{
  return value instanceof Date && !isNaN(value);
};