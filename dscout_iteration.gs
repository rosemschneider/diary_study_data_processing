HEADERS_ROW_INDEX = 1
P_COL_START = 27


function start() {
 // Please update line 7 and 8 to match your sheet data and reflect what you want!
 var oldDataSheetName = 'Test'; // old data sheet of processed data
 var newDataSheetName = 'New_test'; // new data sheet - does not replace your processed data sheet
 const columnsToExclude = ['Scout Name', 'City', 'State', 'Postal Code', 'Education Level', 'Research Industry', 'Entry Headline', 'Time Zone', 'Latitude', 'Longitude', 'User Agent']; //make sure these columns are the same as the ones you initially removed

  //////////////////////////// DO NOT MODIFY ANYTHING BELOW FROM HERE unless you know what you are doing :D ////////// /////////

  // Get data and header for old data
  var oldDataSheetFull = SpreadsheetApp.getActive().getSheetByName(oldDataSheetName);
  var newDataSheetFull = SpreadsheetApp.getActive().getSheetByName(newDataSheetName);
  const oldSheetData = oldDataSheetFull.getDataRange().getValues();
  var newSheetData = newDataSheetFull.getDataRange().getValues();
  const header = [oldSheetData[0]];

  // Get sheet names for old data
  const oldSheetNames = allSheetNames();
  
  // Define the column in which to look for Entry IDs
  var entryIdColumn = 'Entry ID';
  var entryIdColumnIndex = find_column_index_from_column_name(header[0], entryIdColumn)

  // Get the highest entry from the old data seet
  const maxEntry = getHighestEntryId(oldDataSheetFull, entryIdColumnIndex);

   // Make a copy of the sheet for editing
  var tempNewSheet = newDataSheetFull.copyTo(SpreadsheetApp.getActiveSpreadsheet());
  var copiedNewSheetName = "updated_processed_" + newDataSheetName;
  tempNewSheet.setName(copiedNewSheetName);
  var copiedSheet = SpreadsheetApp.getActive().getSheetByName(copiedNewSheetName);
  var copiedData = copiedSheet.getDataRange().getValues();
  const columnsToExcludeInIndex = [];
  const copiedSheetHeader = [copiedData[0]];

  // Sort new data sheet by Entry ID
  newSheetData = sortSheet(copiedSheet, entryIdColumnIndex); //need to update new sheet data

  // get the row with the highest Entry ID
  var highestEntryIdRow = getHighestEntryRow(copiedData, entryIdColumnIndex, maxEntry);

  //on the updated processed data sheet, exclude columns to exclude so that it matches the original sheet
  const columnName = 'Part';
  P_COL_START -= columnsToExclude.length;

//Exclude columns
  for (var i = 0; i < columnsToExclude.length; i++) {
    columnsToExcludeInIndex.push(find_column_index_from_column_name(copiedSheetHeader[0], columnsToExclude[i]));
  }
  columnsToExcludeInIndex.sort(function(a, b) {
    return a - b;
  });


  Logger.log("I will remove " + columnsToExcludeInIndex.length + " columns and these are the columns I will remove: " + columnsToExcludeInIndex);
  for (var j = columnsToExcludeInIndex.length - 1; j >=0; j--) {
    copiedSheet.deleteColumn(columnsToExcludeInIndex[j]);
  }

  // On the new sheet, delete any rows less than or equal to the highest entry ID
  function deleteProcessedData(sheet, highestEntryRow) {
    var numberRowsToDelete = highestEntryRow - 1; 
    sheet.deleteRows(HEADERS_ROW_INDEX + 1, numberRowsToDelete);
  }

  deleteProcessedData(copiedSheet, highestEntryIdRow);

  // Because we have deleted columns on the processed sheet, we now need to redo the process of splitting data
  // we'll copy that new data to the already existing sheets
  // then delete the sheets we've created in the process

  // From here, we want to continue as usual, except that we will be reading in from a different row

    run_group_by(copiedNewSheetName, oldDataSheetName, columnName, oldSheetNames);
  }

  //?
  function find_column_index_from_column_name(header, columnName) {
   for (var i = 0; i < header.length; i++) {
     if (header[i] == columnName) {
       return i+1;
     }
   }
  }

  // Data splitting and merging
  function run_group_by(sheetName, originalDataSheetName, columnName, processedSheetNames) { 
    // get active sheet, data on that sheet, and define where the header is on that sheet
    const sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
    const data = sheet.getDataRange().getValues();
    const header = [data[0]]; 

    // Define the column in which to look for Part values
    const columnIndex = find_column_index_from_column_name(header[0], columnName);

    // Get unique Part values and sort them numerically
    const uniqueValues = getUniqueColumn(sheet, columnIndex);
      uniqueValues.sort(function(a, b) {
      return a - b;
    });

    // Get the lowest Part value
    // Hardcoding this as 1 for now
    const lowest_part_num = 1;

    // Define where we will start scanning data (L->R). This is the LAST column BEFORE data from Parts. 
    // Defining this flexibly because users may sometimes have non-data columns after "Part"

    // need to get the lowest part number - this will need to come from the full dataset    
    const scanning_start_col = get_scanning_start_column(header[0], lowest_part_num);

    // Now start a for loop where we will (by each Part): 
    // 1. Filter data by Part
    // 2. Created a new sheet for that data
    // 3. Get the first column for which we have data. IF there are empty data columns (ie data coming from Parts) before that, bulk delete them. This is the FIRST PASS.
    // 4. After we have anything unnecesary BEFORE the data deleted, we do a SECOND PASS to delete anything unnecessary AFTER the data.
    // 5. ITERATION-specific: We do a final pass to merge the new and old data
     
    // Begin for loop
    for (var i = 0; i < uniqueValues.length; i++) {
      // filter data by part
      const filteredData = data.filter(function(row) {
          return row[columnIndex-1] == uniqueValues[i][0];
      });

      //Create a sheet for data 
      const newSheet = createSheet(sheetName+'_'+getColumnName(sheet, columnIndex)+'='+uniqueValues[i][0]);
      newSheet.getRange(HEADERS_ROW_INDEX, 1, header.length, header[0].length).setValues(header);
      newSheet.getRange(HEADERS_ROW_INDEX + 1, 1, filteredData.length, filteredData[0].length).setValues(filteredData);

      // If we do not have a matching old sheet for this part, then create that sheet as well
      // First, get the part number of the current working sheet 
      var regex = /Part=([0-9]+)/; 
      var value = regex.exec(newSheet.getName())[0];
      var value_str = Utilities.formatString(value);
      Logger.log("The current part number is: " + value); 

      // Second, check to see if we have a matching old sheet
      var matchingSheetName = getMatchingSheet(value_str, processedSheetNames);

      // // Finally, if we do NOT have a matching sheet name, then create one
      // if (matchingSheetName == false) {
      //   Logger.log("I have checked that the matching sheet name is false");
      //   const addedSheet = createSheet("processed_" + originalDataSheetName + getColumnName(sheet, columnIndex)+'='+uniqueValues[i][0]); 
      //   addedSheet.getRange(HEADERS_ROW_INDEX, 1, header.length, header[0].length).setValues(header); // add header
      //   // get the matching sheet name
      //   var matchingSheetName = getMatchingSheet(value_str, processedSheetNames);
      // }
      SpreadsheetApp.flush();

      // BEGIN FIRST PASS - DELETING COLUMNS BEFORE DATA
        //Get the first column for which we have data in this Part
      const first_data_col = get_first_column_part_data(header, uniqueValues[i][0], scanning_start_col);
      // If there are empty data columns before this column, delete them 
      // Empty columns determined by looking at first column after our "scanning start" var and our "first data col" var
       if (first_data_col - (scanning_start_col + 1) != 0) { // If there are empty columns
         // create variables for Google column numbers for bulk deletion because deleteColumns is 1-indexed, but apps script is 0-indexed
         // deleteColumns needs where to start deleting and how many columns to delete after
          var googleScanningStart = scanning_start_col + 2;
          var googleFirstDataCol = first_data_col + 1;
          var googleNumberColsToDelete = googleFirstDataCol - googleScanningStart;
        newSheet.deleteColumns(googleScanningStart, googleNumberColsToDelete); 
      }

      // Alert so we know this has been done
      Logger.log("I am 50% done working on " + newSheet.getSheetName());

      // BEGIN SECOND PASS - DELETING COLUMNS AFTER DATA
      // Because we've had some deletions, we need to update the header so we're indexing appropriately
      const second_pass_data = newSheet.getDataRange().getValues();
      const second_pass_header = [second_pass_data[0]];

      //Now get where we need to start deleting from and how many cols to delete in the second pass
      const second_pass_delete_start = get_first_column_without_data(second_pass_header, uniqueValues[i][0],  scanning_start_col);
      var lastColumn = newSheet.getLastColumn(); 
      
      // now bulk delete columns on the second pass
      if (lastColumn - second_pass_delete_start != 0) {
        // create variables for Google column numbers for bulk deletion because deleteColumns is 1-indexed, but apps script is 0-indexed
        var googleSecondPassDeleteStart = second_pass_delete_start + 1;
        var googleLastCol = lastColumn + 1;
        var googleSecondPassNumberColsDelete = (googleLastCol - googleSecondPassDeleteStart);
        newSheet.deleteColumns(googleSecondPassDeleteStart, googleSecondPassNumberColsDelete);
      }

      // FINAL PASS - MERGING NEW AND OLD DATA
      // First - easy case: When we don't have a matching sheet name, rename current sheet and do nothing else
      if (matchingSheetName == false) {
        SpreadsheetApp.getActive().getSheetByName(sheetName+'_'+getColumnName(sheet, columnIndex)+'='+uniqueValues[i][0]).setName("processed_" + originalDataSheetName + '_' + getColumnName(sheet, columnIndex)+'='+uniqueValues[i][0]);
        Logger.log("I renamed the sheet");
      } else { // if we DO have a matching sheet name

      // First, on the OLD sheet, get the last row on which we have data
      var processedSheet = SpreadsheetApp.getActive().getSheetByName(matchingSheetName);
      var processedSheetLastRow = processedSheet.getLastRow();

         // Now, on the NEW sheet, get the data. We have to update it because of our bulk deletions
      const final_pass_data = newSheet.getRange(HEADERS_ROW_INDEX+1, 1, newSheet.getLastRow(), newSheet.getLastColumn()).getValues();
      // write data from the new sheet to the old sheet
      processedSheet.getRange(processedSheetLastRow + 1, 1, final_pass_data.length, final_pass_data[0].length).setValues(final_pass_data);

      // delete the working sheet
      var ss = SpreadsheetApp.getActive();
      var workingSheet = ss.getSheetByName(sheetName+'_'+getColumnName(sheet, columnIndex)+'='+uniqueValues[i][0]);
      ss.deleteSheet(workingSheet);
      Logger.log("I have written new data to old data sheet and deleted the working sheet");
      }

      //Alert that we're done
      Logger.log("Yay! I am done working on " + value_str);
   }
}

//This is a function for getting the last column before data starts
function get_scanning_start_column(header, min_part_num) {
  var last_initial_column = 0;
  for (var i = 0; i < header.length; i++) {
    if (header[i].indexOf("P"+min_part_num+":") != -1) {
      last_initial_column = i-1;
      break;
    }
  }
  return last_initial_column;
}

//This is a function that gets the first column in which there is data.
function get_first_column_part_data(header, part_num, scanning_start) {
  for (var i = scanning_start+1; i < header[0].length; i++) {
    if (header[0][i].indexOf("P"+part_num+":") != -1) {
      var first_column_with_part_data = i;
      break;
    }
  }
  return first_column_with_part_data; 
}

// This is a function used in the second pass for looking at the END of data to determine where we shoudl start deleting
function get_first_column_without_data(header, part_num, scanning_start) {
  for (var i = scanning_start+1; i < header[0].length; i++) {//now get the first column for which we do not have data after this
   if (header[0][i].indexOf("P"+part_num+":") == -1) {
     var first_col_without_data = i;
     break;
   }
  }
  return first_col_without_data;
}

/*
* Create a new sheet with the given name.
*/
function createSheet(newSheetName) {
   var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
   var yourNewSheet = activeSpreadsheet.getSheetByName(newSheetName);


   if (yourNewSheet != null) {
       activeSpreadsheet.deleteSheet(yourNewSheet);
   }


   yourNewSheet = activeSpreadsheet.insertSheet();
   yourNewSheet.setName(newSheetName);


   Logger.log("I just created sheet " + yourNewSheet.getSheetName() + ". Please do not touch this sheet as I will be working on it.");


   return yourNewSheet;
}

/*
* Helper function for getting the highest entry ID from the old data sheet
*/
  function getHighestEntryId(sheet, column) {
    var entryIdsNumeric = [];
    const rg = sheet.getRange(1 + HEADERS_ROW_INDEX, column, sheet.getLastRow() - HEADERS_ROW_INDEX, 1);
    const entryIds = getUnique1_(rg.getValues()).map(e => [e]);
    for (var i = 0; i < entryIds.length; i++) {
      entryIdsNumeric.push(parseInt(entryIds[i]));
    }
    var maxEntryId = Math.max.apply(Math, entryIdsNumeric);
    Logger.log("The highest Entry ID in the dataset is: '" + maxEntryId);
    return maxEntryId;
  }

  /*
  * Helper function for sorting sheet by a particular column
  */
  function sortSheet(sheet, colIndex) { 
    var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
    rows.sort(colIndex);
    Logger.log("New sheet is sorted by Entry ID");
    var sortedData = sheet.getDataRange().getValues();
    return sortedData;
  }; 


/*
* Helper function for getting the name of the column.
*/
function getColumnName(sheet, column) {
 return sheet.getRange(HEADERS_ROW_INDEX, column).getValue();
}

/*
* Helper function for getting the row corresponding to the highest entry ID
*/
  function getHighestEntryRow(data, colIndex, highestEntryId) {
      for (var i = 0; i < data.length; i++) {
        if(data[i][colIndex-1] == highestEntryId) {
          var highestEntryRow = i + 1;
          break;
        }
      }
    return highestEntryRow;
  }


/*
* Helper function for getting the index of column from column name
*/
 function find_column_index_from_column_name(header, columnName) {
    for (var i = 0; i < header.length; i++) {
      if (header[i] == columnName) {
        return i+1;
      }
    }
  }


/*
* Helper function for getting unique values from 2d arrays.
*/
const getUnique1_ = array2d => [...new Set(array2d.flat())];


/*
* Get a list of unique values from a given column.
*/
function getUniqueColumn(sheet, column) {
 const rg = sheet.getRange(1 + HEADERS_ROW_INDEX, column, sheet.getLastRow() - HEADERS_ROW_INDEX, 1);
 const uniqueValues = getUnique_(rg.getValues()).map(e => [e]);
 Logger.log("I found these parts in your data: '" + getColumnName(sheet, column) + "' => " + uniqueValues);
 return uniqueValues;
}

/*
* Helper function for getting list of sheet names
*/
function allSheetNames() {
  var out = []; 
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  for (var i=0; i<sheets.length; i++) {
    var curSheetName = sheets[i].getName(); 
    if(curSheetName.indexOf('Part') !== -1) {
      out.push(sheets[i].getName()); 
    }
  }
  return out;
}

/*
* Helper function for getting the matching sheet from Part number of active sheet.
* This returns false if there is not a matching sheet name
* If there is not a matching sheet name, then we need to create one
*/ 
function getMatchingSheet(activePartNumberString, processedSheetNames) {
  for (var i=0; i<processedSheetNames.length; i++) {
        if (processedSheetNames[i].indexOf(activePartNumberString) > -1) { 
          var matchingProcessedSheetName = processedSheetNames[i];
          break;
        } else {
          var matchingProcessedSheetName = false;
        }
      }
  Logger.log("The matching processed sheet is: " + matchingProcessedSheetName);
  return matchingProcessedSheetName;
}
