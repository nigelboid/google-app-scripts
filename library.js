/**
 * GetTableByNameSimple()
 *
 * Simplified wrapper for a more flexible GetTableByName()
 */
function GetTableByNameSimple(id, sourceName, verbose)
{
  var firstDataColumn= 0;
  var confirmNumbers= false;
  var limit= null;
  var storeIterationCount= false;
  
  return GetTableByName(id, sourceName, firstDataColumn, confirmNumbers, limit, storeIterationCount, verbose);
};
  
  
/**
 * GetTableByRangeSimple()
 *
 * Simplified wrapper for a more flexible GetTableByRange()
 */
function GetTableByRangeSimple(range, verbose)
{
  var firstDataColumn= 0;
  var confirmNumbers= false;
  var limit= null;
  var storeIterationCount= false;
  
  return GetTableByRange(range, firstDataColumn, confirmNumbers, limit, storeIterationCount, verbose);
};
  
  

/**
 * GetTableByName()
 *
 * Read a table of data into a 2-dimensional array and optionally confirm numeric results
 */
function GetTableByName(id, sourceName, firstDataColumn, confirmNumbers, limit, storeIterationCount, verbose)
{
  var spreadsheet= null;
  var range= null;
  var data= null;
  var table= [];
  var good= true;
  var maxIterations= 10;
  var sleepInterval= 5000;
  var iterationErrors= [];
  
  if (spreadsheet= SpreadsheetApp.openById(id))
  {
    if (range= spreadsheet.getRangeByName(sourceName))
    {
      table= GetTableByRange(range, firstDataColumn, confirmNumbers, limit, storeIterationCount, verbose);
    }
    else
    {
      if (verbose)
      {
        Logger.log("[GetTableByName] Could not get range named <%s> in spreadsheet <%s>.", sourceName, spreadsheet.getName());
      }
      table= null;
    }
  }
  else
  {
    if (verbose)
    {
      Logger.log("[GetTableByName] Could not open spreadsheet ID <%s>.", id);
    }
    table= null;
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
  var data= null;
  var table= [];
  var good= true;
  var maxIterations= 10;
  var sleepInterval= 5000;
  var iterationErrors= [];
  
  for (var iteration= 0; iteration < maxIterations; iteration++)
  {
    iterationErrors= [];
    if (data= range.getValues())
    {
      for (var vIndex= 0; vIndex < data.length; vIndex++)
      {
        if (confirmNumbers)
        {
          for (var hIndex= firstDataColumn; hIndex < data[vIndex].length; hIndex++)
          {
            // check each value obtained to make sure it is a number and above the sanity check limit
            if ((data[vIndex][hIndex] == null) || isNaN(data[vIndex][hIndex]) || (data[vIndex][hIndex] < limit))
            {
              iterationErrors.push("Could not get a viable value (<".concat(data[vIndex][hIndex], "> v. limit of <", limit,
                ">) from location <", hIndex.toFixed(0), ",",  vIndex.toFixed(0), ">."));
              
              data[vIndex][hIndex]= iteration.toFixed(0);
              good= false;
            }
          }
        }
        
        if (good)
        {
          // all values checked out against the limit -- save current data row and append current iteration count (if asked)
          table[vIndex]= data[vIndex];
          if (storeIterationCount)
          {
            table[vIndex].push(iteration.toFixed(0)+1);
          }
        }
      }
      
      if (good)
      {
        // all values checked out against the limit -- we're done here
        break;
      }
      else
      {
        // reset the flag -- we'll try again
        good= true;
      }
    }
    else
    {
      if (verbose)
      {
        Logger.log("[GetTableByRange] Could not read data from range.");
      }
      data= iteration;
    }
    
    Utilities.sleep(sleepInterval);
  }
  
  if (iterationErrors.length > 0)
  {
    // encountered errors while reading data -- report them
    while(iterationErrors.length > 0)
    {
      // report all the accumulated errors
      Logger.log("[GetTableByName] " + iterationErrors.shift());
    }
    Logger.log("[GetTableByName] Reached <%s> iterations but still could not get all data from range.", iteration.toFixed(0));

    table= null;
  }
  
  return table;
};


/**
 * GetValueByName()
 *
 * Obtain a value from a labeled one-cell range
 */
function GetValueByName(id, sourceName, verbose, confirmNumbers, limit)
{
  var value= null;
  var firstDataColumn= 0;
  var storeIterationCount= false;
  
  if (confirmNumbers == undefined)
  {
    confirmNumbers= false;
    limit= 0;
    if (verbose)
    {
      Logger.log("[GetValueByName] confirmNumbers set to default <%s> with limit set to <%s>.", confirmNumbers, limit);
    }
  }
  else
  {
    if (confirmNumbers)
    {
      // make sure limit is defined if we are to confirm numbers
      if (limit == undefined)
      {
        limit= 0;
        if (verbose)
        {
          Logger.log("[GetValueByName] limit set to default <%s>.", limit);
        }
      }
    }
  }
  
  value= GetTableByName(id, sourceName, firstDataColumn, confirmNumbers, limit, storeIterationCount, verbose);
  if (value)
  {
    // We seem to have something!
    if (value.length > 0)
    {
      // We seem to have at least one dimension!
      if (value[0].length > 0)
      {
        // We seem to have a proper table!
        // Return the top-left value
        return value[0][0];
      }
      else
      {
        // Not a proper table!
        if (verbose)
        {
          Logger.log("[GetValueByName] Range named <%s> is not a table.", sourceName);
        }
        return null;
      }
    }
    else
    {
      // Not even a proper array!
      if (verbose)
      {
        Logger.log("[GetValueByName] Range named <%s> is not even an array.", sourceName);
      }
      return null;
    }
  }
  else
  {
    // We got nothing!
    if (verbose)
    {
      Logger.log("[GetValueByName] Range named <%s> did not result in a viable value.", sourceName);
    }
    return null;
  }
};


/**
 * SetTableByName()
 *
 * Write a 2-dimensional array of data into a named spreadsheet table
 */
function SetTableByName(id, destinationName, table, verbose)
{
  var spreadsheet= null;
  var range= null;
  var height= null;
  var width= null;
  var success= true;
  
  if (spreadsheet= SpreadsheetApp.openById(id))
  {
    if (range= spreadsheet.getRangeByName(destinationName))
    {
      // perform capacity checks before writing
      height= range.getHeight();
      if (height < table.length)
      {
        if (verbose)
        {
          Logger.log("[SetTableByName] Could not write out range named <%s> in spreadsheet <%s> since the destiantion is shorter than the data we have by <%s> rows",
                   destinationName, spreadsheet.getName(), table.length - height);
        }
        success= false;
      }
      else
      {
        // looks like we have sufficient height available, now check width
        width= range.getWidth();
        if (width < table[0].length)
        {
          if (verbose)
          {
            Logger.log("[SetTableByName] Could not write out range named <%s> in spreadsheet <%s> since the destiantion is narrower than the data we have by <%s> columns",
                     destinationName, spreadsheet.getName(), table[0].length - width);
          }
          success= false;
        }
        else
        {
          // looks we also have sufficient width available, pad source data (if necessary)
          for (var vIndex= 0; vIndex < height; vIndex++)
          {
            if ((vIndex + 1) > table.length)
            {
              // add blank rows to the source table
              table.push(FillArray(width, ""));
            }
            else if (width > table[vIndex].length)
            {
              // add blank columns to existing rows
              table[vIndex]= table[vIndex].concat(FillArray(width - table[vIndex].length, ""));
            }
          }
          
          // write out the values
          range= range.setValues(table);
          if (!range)
          {
            if (verbose)
            {
              Logger.log("[SetTableByName] Could not write out range named <%s> in spreadsheet <%s>.", destinationName, spreadsheet.getName());
            }
            success= false;
          }
        }
      }
    }
    else
    {
      if (verbose)
      {
        Logger.log("[SetTableByName] Could not get range named <%s> in spreadsheet <%s>.", destinationName, spreadsheet.getName());
      }
      success= false;
    }
  }
  else
  {
    if (verbose)
    {
      Logger.log("[SetTableByName] Could not open spreadsheet ID <%s>.", id);
    }
    success= false;
  }
  
  return success;
};


/**
 * SetValueByName()
 *
 * Set a value from to labeled one-cell range
 */
function SetValueByName(id, destinationName, value, verbose)
{
  return SetTableByName(id, destinationName, [[value]], verbose)
};


/**
 * GetAnnualSheetIDs()
 *
 * Looks up IDs of all known annual sheets
 */
function GetAnnualSheetIDs(id, verbose)
{
  var sourceName= "ExternalLookups";
  var idsByYear= [];
  var sheetIDs= {};
  
  idsByYear= GetTableByNameSimple(id, sourceName, verbose);
  
  if (idsByYear)
  {
    // we have viable IDs
    for (var vIndex= 0; vIndex < idsByYear.length; vIndex++)
    {
      sheetIDs[idsByYear[vIndex][0]]= idsByYear[vIndex][1];
    }
  }
  else
  {
    if (verbose)
    {
      Logger.log("[GetAnnualSheets] Could not obtain a list of annual sheet IDs from table <%s> of spreadsheet ID <%s>.", sourceName, id);
    }
  }
  
  return sheetIDs;
};


/**
 * SaveValue()
 *
 * Save current values in a mirror table
 */
function SaveValue(id, sourceName, destinationName, verbose, confirmNumbers, limit)
{
  var sourceValues= [];
  var destinationValues= [];
  var firstDataColumn= 0;
  var storeIterationCount= false;
  var changed= false;
  
  // set defaults unless supplied
  if (confirmNumbers == undefined)
  {
    confirmNumbers= false;
    limit= 0;
    if (verbose)
    {
      Logger.log("[SaveValue] confirmNumbers set to default <%s> with limit set to <%s>.", confirmNumbers, limit);
    }
  }
  else
  {
    if (confirmNumbers)
    {
      // make sure limit is defined if we are to confirm numbers
      if (limit == undefined)
      {
        limit= 0;
        if (verbose)
        {
          Logger.log("[SaveValue] limit set to default <%s>.", limit);
        }
      }
    }
  }
  
  // read all the source and destination values, compare, and update
  if (sourceValues= GetTableByName(id, sourceName, firstDataColumn, confirmNumbers, limit, storeIterationCount, verbose))
  {
    // we have source values, proceed to destination values
    if (destinationValues= GetTableByName(id, destinationName, firstDataColumn, confirmNumbers, limit, storeIterationCount, verbose))
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
                if (verbose)
                {
                  Logger.log("[SaveValue] Value at location <%s, %s> has changed to <%s> in table <%s> from <%s> in table <%s> of spreadsheet ID <%s>.",
                             hIndex.toFixed(0), vIndex.toFixed(0), sourceValues[vIndex][hIndex], sourceName, destinationValues[vIndex][hIndex], destinationName, id);
                }
                destinationValues[vIndex][hIndex]= sourceValues[vIndex][hIndex];
                changed= true;
              }
              else if (verbose)
              {
                Logger.log("[SaveValue] Value <%s> (<%s>) at location <%s, %s> has not changed between named tables <%s> and <%s> of spreadsheet ID <%s>.",
                           destinationValues[vIndex][hIndex], sourceValues[vIndex][hIndex], hIndex.toFixed(0), vIndex.toFixed(0), sourceName, destinationName, id);
              }
            }
          }
          else
          {
            Logger.log("[SaveValue] Source values range <%s, %s> of source <%s> does not match destination range <%s, %s> of destination <%s> in sheet ID <%s>.",
                       sourceValues.length, sourceValues[vIndex].length, sourceName, destinationValues.length, destinationValues[vIndex].length, destinationName, id);
          }
        }
      }
      else
      {
        Logger.log("[SaveValue] Source values height <%s> of source <%s> does not match destination range <%s> of destination <%s> in sheet ID <%s>.",
                   sourceValues.length, sourceName, destinationValues.length, destinationName, id);
      }
      
      if (changed)
      {
        // write out the values
        if (!SetTableByName(id, destinationName, destinationValues, verbose))
        {
          // something went wrong!
          Logger.log("[SaveValue] Could not write out range named <%s> in spreadsheet ID <%s>.", destinationName, id);
          
          changed= false;
        }
      }
    }
    else
    {
      Logger.log("[SaveValue] Could not get range named <%s> in spreadsheet ID <%s>.", destinationName, id);
    }
  }
  else
  {
    Logger.log("[SaveValue] Could not get range named <%s> in spreadsheet ID <%s>.", sourceName, id);
  }
  
  return changed;
};


/**
 * GetLastSnapshotStamp()
 *
 * Obtain the identifier stamp for the last snapshot entry
 */
function GetLastSnapshotStamp(id, sheetName, verbose)
{
  var spreadsheet= null;
  var sheet= null;
  var range= null;
  var height= null;
  var value= null;
  
  if (spreadsheet= SpreadsheetApp.openById(id))
  {
    if (sheet= spreadsheet.getSheetByName(sheetName))
    {
      if (height= sheet.getLastRow())
      {
        if (range= sheet.getRange(height, 1))
        {
          value= range.getValue();
        }
        else if (verbose)
        {
          Logger.log("[GetLastSnapshotStamp] Could not set range to the first cell of the last row <%s> in sheet <%s> for spreadsheet <%s>.",
                     height.toFixed(0), sheetName, spreadsheet.getName());
        }
      }
      else if (verbose)
      {
        Logger.log("[GetLastSnapshotStamp] Could not learn the last row in sheet <%s> for spreadsheet <%s>.", sheetName, spreadsheet.getName());
      }
    }
    else if (verbose)
    {
      Logger.log("[GetLastSnapshotStamp] Could not activate sheet <%s> for spreadsheet <%s>.", sheetName, spreadsheet.getName());
    }
  }
  else if (verbose)
  {
    Logger.log("[GetLastSnapshotStamp] Could not open spreadsheet ID <%s>.", id);
  }
  
  return value;
};


/**
 * SelectCell()
 *
 * Obtain a range specification for a given cell of the named sheet
 */
function SelectCell(id, sheetName, cellCoordinates, verbose)
{
  var spreadsheet= null;
  var sheet= null;
  var range= null;
  
  if (spreadsheet= SpreadsheetApp.openById(id))
  {
    if (sheet= spreadsheet.getSheetByName(sheetName))
    {
      range= sheet.getRange(cellCoordinates);
    }
    else
    {
      Logger.log("[SelectCell] Could not activate sheet <%s> for spreadsheet <%s>.", sheetName, spreadsheet.getName());
    }
  }
  else
  {
    Logger.log("[SelectCell] Could not open spreadsheet ID <%s>.", id);
  }

  return range;
};


/**
 * GetCellValue()
 *
 * Obtain the value of the specified cell
 */
function GetCellValue(id, sheetName, cellCoordinates, verbose)
{
  var range= SelectCell(id, sheetName, cellCoordinates, verbose);
  var value= null;
  
  if (range)
  {
    value= range.getValue();
  }
  else if (verbose)
  {
    Logger.log("[GetCellValue] Could not access specified range <%s> in sheet <%s> for spreadsheet ID <%s>.",
                cellCoordinates, sheetName, id);
  }
  
  return value;
};


/**
 * SetCellValue()
 *
 * Set the value of the specified cell
 */
function SetCellValue(id, sheetName, cellCoordinates, value, verbose)
{
  var range= SelectCell(id, sheetName, cellCoordinates, verbose);
  var success= false;
  
  if (range)
  {
    success= range.setValue(value);
  }
  else if (verbose)
  {
    Logger.log("[SetCellValue] Could not access specified range <%s> in sheet <%s> for spreadsheet ID <%s>.", cellCoordinates, sheetName, id);
  }
  
  return success;
};


/**
 * CheckSnapshot()
 *
 * Check data in the destination spreadsheet
 */
function CheckSnapshot(id, sheetName, newDataDate, verbose)
{
  var lastDataDate= new Date(GetLastSnapshotStamp(id, sheetName, verbose));
  
  if (lastDataDate && (lastDataDate.getFullYear() == newDataDate.getFullYear()) &&
     (lastDataDate.getMonth() == newDataDate.getMonth()) && (lastDataDate.getDate() == newDataDate.getDate()))
  {
    // we have already recorded data for today
    return true;
  }
  else if (lastDataDate > newDataDate)
  {
    Logger.log("[CheckSnapshot] We seem to have stale data from the past (last date <%s> is later than new date <%s>), skipping...",
               lastDataDate, newDataDate);
    
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
function CompileSnapshot(id, names, limits, now, verbose)
{
  var snapshot= [now];
  var table= [];
  var firstDataColumn= 0;
  var confirmNumbers= false;
  var storeIterationCount= true;
  var good= true;
  var iterations= 0;
  
  for (var counter= 0; counter < names.length; counter++)
  {
    // check whether we should confirm numeric values for this data point
    if (Number.isNaN(limits[counter]) || typeof limits[counter] != "number")
    {
      confirmNumbers= false;
    }
    else
    {
      confirmNumbers= true;
    }

    // read each table or cell of interest and accumulate in an array
    table= GetTableByName(id, names[counter], firstDataColumn, confirmNumbers, limits[counter], storeIterationCount, verbose);
    if (table)
    {
      // we got viable data, now transpose the table
      for (var row= 0; row < table.length; row++)
      {
        // grab the first column value from every row returned
        snapshot.push(table[row][0]);
        if (iterations < table[row][table[row].length-1])
        {
          // store the highest iteration count
         iterations= table[row][table[row].length-1];
        }
      }
    }
    else
    {
      // we failed to obtain viable data
      good= false;

      Logger.log("[CompileSnapshot] Failed to read data for <%s> from sheet id <%s>, skipping the entire snapshot...", names[counter], id);

      break;
    }
  }
  
  if (good)
  {
    // dress and return viable data
    snapshot[0]= new Date();
    snapshot.push(iterations);
    return snapshot;
  }
  else
  {
    // no viable data to return
    return null;
  }
};


/**
 * SaveSnapshot()
 *
 * Save values snapshot in a history table
 */
function SaveSnapshot(id, sheetName, values, updateRun, verbose)
{
  var spreadsheet= null;
  var sheet= null;
  var range= null;
  var lastRow= null;
  var success= false;
  
  if (values)
  {
    // we have viable data to save
    if (!Array.isArray(values[0]))
    {
      // we seem to have a one-dimensional array -- convert it
      values= [values];
    }
    
    // now access the spreadsheet and save
    if (spreadsheet= SpreadsheetApp.openById(id))
    {
      if (sheet= spreadsheet.getSheetByName(sheetName))
      {
        if (lastRow= sheet.getLastRow())
        {
          if (!updateRun)
          {
            // this is not an update run -- append a row
            lastRow++;
          }
          
          if (range= sheet.getRange(lastRow, 1, values.length, values[0].length))
          {
            if (range= range.setValues(values))
            {
              if (PropagateFormulas(sheet, lastRow, values[0].length, verbose))
              {
                success= true;
              }
              else if (verbose)
              {
                Logger.log("[SaveSnapshot] Could not propagate formulas in sheet <%s> for spreadsheet <%s>.", sheetName, spreadsheet.getName());
              }
            }
            else if (verbose)
            {
              Logger.log("[SaveSnapshot] Could not append values in sheet <%s> for spreadsheet <%s>.", sheetName, spreadsheet.getName());
            }
          }
          else if (verbose)
          {
            Logger.log("[SaveSnapshot] Could not set range to append beyond the last row <%s> in sheet <%s> for spreadsheet <%s>.",
                       height.toFixed(0), sheetName, spreadsheet.getName());
          }
        }
        else if (verbose)
        {
          Logger.log("[SaveSnapshot] Could not learn the last row in sheet <%s> for spreadsheet <%s>.", sheetName, spreadsheet.getName());
        }
      }
      else if (verbose)
      {
        Logger.log("[SaveSnapshot] Could not activate sheet <%s> for spreadsheet <%s>.", sheetName, spreadsheet.getName());
      }
    }
    else if (verbose)
    {
      Logger.log("[SaveSnapshot] Could not open spreadsheet ID <%s>.", id);
    }
  }
  else if (verbose)
  {
    Logger.log("[SaveSnapshot] Nothing to write to sheet <%s> of spreadsheet ID <%s>...", sheetName, id);
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
  var width= null;
  var formulas= null;
  var range= null;
  var success= false;
  
  if (width= sheet.getLastColumn())
  {
    if (width > column)
    {
      // looks like we have spare columns to check
      if (formulas= sheet.getRange(row-1, column+1, 1, width-column).getFormulas())
      {
        if (success= sheet.getRange(row, column+1, 1, width-column).setFormulas(formulas))
        {
          if (verbose)
          {
            Logger.log("[PropagateFormulas] Updated formulas in columns <%s> through <%s> of row <%s> in sheet <%s>.",
                       (column+1).toFixed(0), width.toFixed(0), row.toFixed(0), sheet.getName());
          }
        }
        else if (verbose)
        {
          Logger.log("[PropagateFormulas] Could not set formulas in columns <%s> through <%s> of row <%s> in sheet <%s>.",
                     (column+1).toFixed(0), width.toFixed(0), row.toFixed(0), sheet.getName());
        }
      }
      else if (verbose)
      {
        Logger.log("[PropagateFormulas] Could not read formulas from columns <%s> through <%s> of row <%s> in sheet <%s>.",
                   (column+1).toFixed(0), width.toFixed(0), (row-1).toFixed(0), sheet.getName());
      }
    }
    else
    {
      if (verbose)
      {
        Logger.log("[PropagateFormulas] No columns to propagate in sheet <%s>.", sheet.getName());
      }
      success= true;
    }
  }
  else if (verbose)
  {
    Logger.log("[PropagateFormulas] Could not obtain width of sheet <%s>.", sheet.getName());
  }
  
  return success;
};


/**
 * UpdateSnapshotCell()
 *
 * Update a specific value in a history table
 */
function UpdateSnapshotCell(id, sheetName, column, value, onlyIfBlank, verbose)
{
  var spreadsheet= null;
  var sheet= null;
  var range= null;
  var height= null;
  var success= false;
  
  if (value != null)
  {
    // now access the spreadsheet and save
    if (spreadsheet= SpreadsheetApp.openById(id))
    {
      if (sheet= spreadsheet.getSheetByName(sheetName))
      {
        if (height= sheet.getLastRow())
        {
          if (range= sheet.getRange(height, column, 1, 1))
          {
            if (!onlyIfBlank || range.isBlank())
            {
              if (range= range.setValue([[value]]))
              {
                if (verbose)
                {
                  Logger.log("[UpdateSnapshotCell] Updated cell <%s> of the last row <%s> in sheet <%s> for spreadsheet <%s> with <%s>.",
                             column.toFixed(0), height.toFixed(0), sheetName, spreadsheet.getName(), value);
                }
                success= true;
              }
              else if (verbose)
              {
                Logger.log("[UpdateSnapshotCell] Could not update cell <%s> of the last row <%s> in sheet <%s> for spreadsheet <%s>.",
                           column.toFixed(0), height.toFixed(0), sheetName, spreadsheet.getName());
              }
            }
            else if (verbose)
            {
              Logger.log("[UpdateSnapshotCell] Could not update cell <%s> of the last row <%s> in sheet <%s> for spreadsheet <%s> "
                         + "since that would clobber an existing value <%s>.", column.toFixed(0), height.toFixed(0), sheetName, spreadsheet.getName(), range.getValue());
            }
          }
          else if (verbose)
          {
            Logger.log("[UpdateSnapshotCell] Could not set range to update the last row <%s> in sheet <%s> for spreadsheet <%s>.",
                       height.toFixed(0), sheetName, spreadsheet.getName());
          }
        }
        else if (verbose)
        {
          Logger.log("[UpdateSnapshotCell] Could not learn the last row in sheet <%s> for spreadsheet <%s>.", sheetName, spreadsheet.getName());
        }
      }
      else if (verbose)
      {
        Logger.log("[UpdateSnapshotCell] Could not activate sheet <%s> for spreadsheet <%s>.", sheetName, spreadsheet.getName());
      }
    }
    else if (verbose)
    {
      Logger.log("[UpdateSnapshotCell] Could not open spreadsheet ID <%s>.", id);
    }
  }
  else if (verbose)
  {
    Logger.log("[UpdateSnapshotCell] Nothing to update in column <%s> of sheet <%s> in spreadsheet ID <%s>...", column, sheetName, id);
  }
  
  return success;
};


/**
 * SaveValuesInHistory()
 *
 * Saves current values in a history table
 */
function SaveValuesInHistory(id, sheetName, sourceNames, sourceLimits, now, backupRun, updateRun, verbose)
{
  if (CheckSnapshot(id, sheetName, now, verbose))
  {
    if (updateRun)
    {
      SaveSnapshot(id, sheetName, CompileSnapshot(id, sourceNames, sourceLimits, now, verbose), updateRun, verbose);
    }
    else if (!backupRun)
    {
      Logger.log("[SaveValuesInHistory] Redundant primary run at <%s> for sheet <%s> in spreadhseet ID <%s>.", now, sheetName, id);
    }
  }
  else
  {
    if (backupRun)
    {
      Logger.log("[SaveValuesInHistory] Primary run seems to have failed for sheet <%s> in spreadhseet ID <%s>...", sheetName, id);
    }
    
    SaveSnapshot(id, sheetName, CompileSnapshot(id, sourceNames, sourceLimits, now, verbose), updateRun= false, verbose);
  }
};


/**
 * RemoveDuplicateSnapshot()
 *
 * Remove a recent duplicate entry from the history table
 */
function RemoveDuplicateSnapshot(id, sheetName, verbose)
{
  var spreadsheet= null;
  var sheet= null;
  var range= null;
  var height= null;
  var width= null;
  var rowData= null;
  var ultimateStamp= null;
  var penultimateStamp= null;
  var priorStamp= null;
  
  if (spreadsheet= SpreadsheetApp.openById(id))
  {
    if (sheet= spreadsheet.getSheetByName(sheetName))
    {
      if (height= sheet.getLastRow())
      {
        // Learn the latest time stamp
        if (range= sheet.getRange(height, 1))
        {
          ultimateStamp= new Date(range.getValue());
        }
        else
        {
          Logger.log("[RemoveDuplicateSnapshot] Could not set range to the first cell of the last row <%s> in sheet <%s> "
                      + "for spreadsheet <%s>.", height.toFixed(0), sheetName, spreadsheet.getName());
        }
        
        
        // Learn the prior time stamp (seemingly real and accurate data two rows above latest)
        if (range= sheet.getRange(height-2, 1))
        {
          priorStamp= new Date(range.getValue());
        }
        else
        {
          Logger.log("[RemoveDuplicateSnapshot] Could not set range to the first cell of the prior row <%s> in sheet <%s> "
                      + "for spreadsheet <%s>.", (height-2).toFixed(0), sheetName, spreadsheet.getName());
        }
        
        // Learn the time stamp just before the latest
        if (width= sheet.getLastColumn())
        {
          if (range= sheet.getRange(height-1, 1, 1, width))
          {
            rowData= range.getValues();
            penultimateStamp= new Date(rowData[0][0]);
          }
          else
          {
            Logger.log("[RemoveDuplicateSnapshot] Could not set range to the first cell of the next to the last row <%s> in sheet <%s> "
                        + "for spreadsheet <%s>.", (height-1).toFixed(0), sheetName, spreadsheet.getName());
          }
        }
        else
        {
          Logger.log("[RemoveDuplicateSnapshot] Could not learn the last column in sheet <%s> for spreadsheet <%s>.",
                      sheetName, spreadsheet.getName());
        }
        
        // Remove the next to the last row if its time stamp matches that of the last row or the prior row preceding it
        if (ultimateStamp.getTime() === penultimateStamp.getTime() || priorStamp.getTime() === penultimateStamp.getTime())
        {
          try
          {
            sheet= sheet.deleteRow(height-1)
          }
          catch (error)
          {
            Logger.log("[RemoveDuplicateSnapshot] Failed to remove duplicate row:\n".concat(error));
          }
          
          if (sheet)
          {
            return [[priorStamp, "Kept"], rowData[0], [ultimateStamp, "Kept"]];
          }
          else
          {
            Logger.log("[RemoveDuplicateSnapshot] Failed to remove the penultimate row for time stamp <%s> in sheet <%s> for spreadsheet <%s>.",
                       penultimateStamp, sheetName, spreadsheet.getName());
          }
        }
        else
        {
          if (verbose)
          {
            Logger.log("[RemoveDuplicateSnapshot] No need to remove history rows as time stamps (<%s> and <%s>) do not match in sheet <%s> for spreadsheet <%s>.",
                       ultimateStamp, penultimateStamp, sheetName, spreadsheet.getName());
          }
        }
      }
      else
      {
        Logger.log("[RemoveDuplicateSnapshot] Could not learn the last row in sheet <%s> for spreadsheet <%s>.", sheetName, spreadsheet.getName());
      }
    }
    else
    {
      Logger.log("[RemoveDuplicateSnapshot] Could not activate sheet <%s> for spreadsheet <%s>.", sheetName, spreadsheet.getName());
    }
  }
  else
  {
    Logger.log("[RemoveDuplicateSnapshot] Could not open spreadsheet ID <%s>.", id);
  }
  
  return false;
};


/**
 * TrimHistory()
 *
 * Remove earliest entries from the history table
 */
function TrimHistory(id, sheetName, maxRows, verbose)
{
  var spreadsheet= null;
  var sheet= null;
  var height= 0;
  
  if (spreadsheet= SpreadsheetApp.openById(id))
  {
    if (sheet= spreadsheet.getSheetByName(sheetName))
    {
      height= sheet.getLastRow();
      
      // Accommodate the header row!
      if (height > (maxRows + 1))
      {
        try
        {
          sheet.deleteRows(2, height - (maxRows + 1));
        }
        catch (error)
        {
          Logger.log("[TrimHistory] Failed to trim history rows:\n".concat(error));
        }
      }
    }
    else
    {
      Logger.log("[TrimHistory] Could not activate sheet <%s> for spreadsheet <%s>.", sheetName, spreadsheet.getName());
    }
  }
  else
  {
    Logger.log("[TrimHistory] Could not open spreadsheet ID <%s>.", id);
  }
};


/**
 * Synchronize()
 *
 * Preserve value from one range in another range (designed to work on first element of each range)
 */
function Synchronize(sourceID, destinationID, sourceNames, destinationNames, verbose, verboseChanges, numbersOnly)
{
  var spreadsheet= null;
  var sheet= null;
  var range= null;
  var value= null;
  var format= null;
  var sourceValues= [];
  var success= true;
  
  if (verboseChanges == undefined)
  {
    // initialize omitted optional verbose flag to match overall verbose flag
    verboseChanges= verbose;
  }
  
  if (numbersOnly == undefined)
  {
    // initialize omitted optional verbose flag to match overall verbose flag
    numbersOnly= false;
  }
  
  if (spreadsheet= SpreadsheetApp.openById(sourceID))
  {
    for (var counter= 0; counter < sourceNames.length; counter++)
    {
      // read all the source values
      if (range= spreadsheet.getRangeByName(sourceNames[counter]))
      {
        value= range.getValue();
        format= range.getNumberFormat();
        if (Number.isNaN(value) || typeof value != "number")
        {
          if (verbose || numbersOnly)
          {
            Logger.log("[Synchronize] Failed to obtain numeric value <%s> with format <%s> from range named <%s> in spreadsheet <%s>.",
                       value, format, sourceNames[counter], spreadsheet.getName());
          }
          
          if (numbersOnly)
          {
            // fail gently
            
            sourceValues.push(null);
            success= false;
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
            // round all currency amounts to cents
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
        if (verbose)
        {
          Logger.log("[Synchronize] Could not get range named <%s> in spreadsheet <%s>.", sourceNames[counter], spreadsheet.getName());
        }
        success= false;
      }
    }
  }
  else
  {
    if (verbose)
    {
      Logger.log("[Synchronize] Could not open spreadsheet ID <%s>.", sourceID);
    }
    success= false;
  }
  
  if (sourceValues.length == destinationNames.length)
  {
    if (spreadsheet= SpreadsheetApp.openById(destinationID))
    {
      for (var counter= 0; counter < destinationNames.length; counter++)
      {
        // compare and write values
        if (range= spreadsheet.getRangeByName(destinationNames[counter]))
        {
          value= range.getValue();
          
          if ((sourceValues[counter] != null) && (value != sourceValues[counter]))
          {
            // looks like the value has changed -- update it
            range.setValue(sourceValues[counter]);
            if (verboseChanges)
            {
              Logger.log("[Synchronize] Value for range <%s> in sheet <%s> updated to <%s>, it was <%s>.",
                         destinationNames[counter], spreadsheet.getName(), sourceValues[counter], value);
            }
          }
          else
          {
            if (verbose)
            {
              Logger.log("[Synchronize] Value for range <%s> has not changed <%s>.", destinationNames[counter], sourceValues[counter]);
            }
          }
        }
        else
        {
          if (verbose)
          {
            Logger.log("[Synchronize] Could not get range named <%s> in spreadsheet <%s>.", destinationNames[counter], spreadsheet.getName());
          }
          success= false;
        }
      }
    }
    else
    {
      if (verbose)
      {
        Logger.log("[Synchronize] Could not open spreadsheet ID <%s>.", destinationID);
      }
      success= false;
    }
  }
  else
  {
    if (verbose)
    {
      Logger.log("[Synchronize] Source values range <%s> does not match destination range <%s>.", sourceValues.length.toFixed(0), destinationNames.length.toFixed(0));
    }
    success= false;
  }
  
  return success;
};


/**
 * GetParameters()
 *
 * Read specified table and return an associative array comprised of key-value pairs from the first two columns
 */
function GetParameters(id, sourceName, verbose)
{
  var parameters= {"id": id, "verbose": verbose};
  var firstDataColumn= 1;
  var confirmNumbers= false;
  var limit= 0;
  var storeIterationCount= false;
  var table= GetTableByName(id, sourceName, firstDataColumn, confirmNumbers, limit, storeIterationCount, verbose);
  
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
            parameters[table[row][0]]= table[row][1];
            parameters[table[row][0].toLowerCase()]= table[row][1];
          }
        }
      }
      else if (verbose)
      {
        // Not a proper table!
        Logger.log("[GetParameters] Range named <%s> is not a table.", sourceName);
      }
    }
    else if (verbose)
    {
      // Not even a proper array!
      Logger.log("[GetParameters] Range named <%s> is not even an array.", sourceName);
    }
  }
  else if (verbose)
  {
    // We got nothing!
    Logger.log("[GetParameters] Range named <%s> did not result in a viable value.", sourceName);
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
  var fill= [];
  var counter= 0;
  
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
  var formattedNumber= "" + number;
  
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
  var specificaion= null;
  const firstColumn= "A";
      
  if (column > ("Z".charCodeAt(0) - firstColumn.charCodeAt(0) + 1))
  {
    if (verbose)
    {
      Logger.log("[SpecifyColumnA1Notation] No support (yet) for managing a column <%s> beyond position 'Z.'", column);
    }
  }
  else
  {
    specificaion= String.fromCharCode(firstColumn.charCodeAt(0) + column);
  }
  
  return specificaion;
};


/**
 * DateToLocaleString()
 *
 * Return a formatted date as a short string (m/dd/yyyy hh:mm:ss)
 */
function DateToLocaleString(date, separator)
{
  var dateOptions= { day: '2-digit', month: '2-digit', year: 'numeric' };
  //var timeOptions= { hour12: false, hourCycle: 'h23', hour: '2-digit', minute:'2-digit', second: '2-digit'};
  var timeOptions= { hourCycle: 'h23', hour: '2-digit', minute:'2-digit', second: '2-digit'};
  
  if (date == undefined)
  {
    date= new Date();
  }
  else
  {
    date= new Date(date);
  }
  
  if (separator == undefined)
  {
    separator= " ";
  }
  
  //return date.toLocaleString('en-US', {hour12: false, hourCycle: 'h23'});
  //return date.toLocaleString('en-US', {hourCycle: 'h23'});
  return date.toLocaleDateString('en-US', dateOptions) + separator + date.toLocaleTimeString('en-US', timeOptions)
};


/**
 * DumpObjectAsPrettyText()
 *
 * Return a formatted, indented text block of object data
 */
function DumpObjectAsPrettyText(data, indentation, level)
{
  var result= "";
  var actualIndentation= "";
  
  if (indentation == null)
  {
    indentation= "  ";
  }
  
  if (level == null || level == 0)
  {
    level= 1;
  }
  
  for (var iteration= 0; iteration < level; iteration++)
  {
    actualIndentation+= indentation;
  }

  for (const property in data)
  {
    var element= data[property];
    var bracketLeft= "{";
    var bracketRight= "}";
    
    if (typeof element == 'string')
    {
      element= "'" + element + "'";
    }
    else if (typeof element == 'object')
    {
      if (element instanceof Array)
      {
        bracketLeft= "[";
        bracketRight= "]";
      }
      
      if (element instanceof Date)
      {
        element= "\n" + actualIndentation + bracketLeft + DateToLocaleString(element) + bracketRight;
      }
      else
      {
        // Recursive dump
        var subSection= DumpObjectAsPrettyText(element, indentation, level + 1);
        
        //element= "\n" + actualIndentation + "( " + Object.getPrototypeOf(element) + " ) " + bracketLeft + "\n" + subSection + "\n" + actualIndentation + bracketRight;
        element= "\n" + actualIndentation + bracketLeft + "\n" + subSection + "\n" + actualIndentation + bracketRight;
      }
    }
    
    result+= actualIndentation + "'" + property + "' : " + element + ",\n";
  }
  
  return result.replace(/,\n$/, "");
}