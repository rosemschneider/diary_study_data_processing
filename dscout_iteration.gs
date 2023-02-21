HEADERS_ROW_INDEX = 1
P_COL_START = 27


function start() {
 // Please update line 7 and 8 to match your sheet data and reflect what you want!
 var oldSheetName = 'Test'; // old data sheet of processed data
 var newSheetName = 'New_test'; // new data sheet - does not replace your processed data sheet
 const columnsToExclude = ['Scout Name', 'City', 'State', 'Postal Code', 'Education Level', 'Research Industry', 'Entry Headline', 'Time Zone', 'Latitude', 'Longitude', 'User Agent']; //make sure these columns are the same as the ones you initially removed


  //////////////////////////// DO NOT MODIFY ANYTHING BELOW FROM HERE unless you know what you are doing :D ////////// /////////

  // Get data and header for old data
  var oldSheet = SpreadsheetApp.getActive().getSheetByName(oldSheetName);
  var newSheet = SpreadsheetApp.getActive().getSheetByName(newSheetName);
  const oldSheetData = oldSheet.getDataRange().getValues();
  var newSheetData = newSheet.getDataRange().getValues();
  const header = [oldSheetData[0]];

  // Define the column in which to look for Entry IDs
  var entryIdColumn = 'Entry ID';
  var entryIdColumnIndex = find_column_index_from_column_name(header[0], entryIdColumn)
  Logger.log("This is the EntryId Column index: " + entryIdColumnIndex);

  // Get the highest entry ID from the old datasheet
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

  const maxEntry = getHighestEntryId(oldSheet, entryIdColumnIndex);

  function find_column_index_from_column_name(header, columnName) {
    for (var i = 0; i < header.length; i++) {
      if (header[i] == columnName) {
        return i+1;
      }
    }
  }

   // Make a copy of the sheet for editing
  var tempNewSheet = newSheet.copyTo(SpreadsheetApp.getActiveSpreadsheet());
  var copiedNewSheetName = "updated_processed_" + newSheetName;
  var copiedSheet = SpreadsheetApp.getActive().getSheetByName(copiedNewSheetName);
  var copiedData = copiedSheet.getDataRange().getValues();

  tempNewSheet.setName(copiedNewSheetName);
  
  // Sort copied sheet based on Entry ID column
  function sortSheet(sheet, colIndex) { 
    var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
    rows.sort(colIndex);
    Logger.log("New sheet is sorted by Entry ID");
    var sortedData = sheet.getDataRange().getValues();
    return sortedData;
  }; 

  newSheetData = sortSheet(copiedSheet, entryIdColumnIndex); //need to update new sheet data

  // Get the row with the highest entry ID
  function getHighestEntryRow(data, colIndex, highestEntryId) {
      for (var i = 0; i < data.length; i++) {
        if(data[i][colIndex-1] == highestEntryId) {
          var highestEntryRow = i + 1;
          break;
        }
      }
    return highestEntryRow;
  }

  var highestEntryIdRow = getHighestEntryRow(copiedData, entryIdColumnIndex, maxEntry);

  // // // Delete any rows less than or equal to the highest entry ID
  // function deleteProcessedData(sheet, highestEntryRow) {
  //   var numberRowsToDelete = highestEntryRow - 1; 
  //   sheet.deleteRows(2, numberRowsToDelete);
  // }

  // deleteProcessedData(copiedNewSheetName, highestEntryIdRow);

  // From here, we likely just want to continue as usual, except that we will be reading in from a different row



  // const columnName = 'Part';
  // P_COL_START -= columnsToExclude.length;


  // // Make a copy of the original sheet with excluded columns removed. 
  // // Additional data processing run on the copied sheet.
  //  var sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
  //  var tempOriginalSheet = sheet.copyTo(SpreadsheetApp.getActiveSpreadsheet());
  //  tempOriginalSheet.setName("processed_" + sheetName);


  //  sheetName = "processed_" + sheetName;
  //  sheet = tempOriginalSheet;


  //  const columnsToExcludeInIndex = []
  //  const data = sheet.getDataRange().getValues();
  //  const header = [data[0]];

  // //Exclude columns
  //  for (var i = 0; i < columnsToExclude.length; i++) {
  //    columnsToExcludeInIndex.push(find_column_index_from_column_name(header[0], columnsToExclude[i]));
  //  }
  //   columnsToExcludeInIndex.sort(function(a, b) {
  //    return a - b;
  //  });


  //  Logger.log("I will remove " + columnsToExcludeInIndex.length + " columns and these are the columns I will remove: " + columnsToExcludeInIndex);
  //  for (var j = columnsToExcludeInIndex.length - 1; j >=0; j--) {
  //    sheet.deleteColumn(columnsToExcludeInIndex[j]);
  //  }
  //   run_group_by(sheetName, columnName);
  // }


  // function find_column_index_from_column_name(header, columnName) {
  //  for (var i = 0; i < header.length; i++) {
  //    if (header[i] == columnName) {
  //      return i+1;
  //    }
  //  }
  // }

  // This function does the following: 
  // 1. Gets unique part values
  // 2. Filters data 
  // function run_group_by(sheetName, columnName) {
  //   // get active sheet, data on that sheet, and define where the header is on that sheet
  //   const sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
  //   const data = sheet.getDataRange().getValues();
  //   const header = [data[0]]; 

  //   // Define the column in which to look for Part values
  //   const columnIndex = find_column_index_from_column_name(header[0], columnName)

  //   // Get unique Part values and sort them numerically
  //   const uniqueValues = getUniqueColumn(sheet, columnIndex);
  //     uniqueValues.sort(function(a, b) {
  //     return a - b;
  //   });

  //   // Get the lowest Part value
  //   const lowest_part_num = Math.min(parseInt(uniqueValues));

  //   // Define where we will start scanning data (L->R). This is the LAST column BEFORE data from Parts. 
  //   // Defining this flexibly because users may sometimes have non-data columns after "Part"
  //   const scanning_start_col = get_scanning_start_column(header[0], lowest_part_num);

  //   // Now start a for loop where we will (by each Part): 
  //   // 1. Filter data by Part
  //   // 2. Created a new sheet for that data
  //   // 3. Get the first column for which we have data. IF there are empty data columns (ie data coming from Parts) before that, bulk delete them. This is the FIRST PASS.
  //   // 4. After we have anything unnecesary BEFORE the data deleted, we do a SECOND PASS to delete anything unnecessary AFTER the data.

  //   // Begin for loop
  //   for (var i = 0; i < uniqueValues.length; i++) {
  //     // filter data by part
  //     const filteredData = data.filter(function(row) {
  //         return row[columnIndex-1] == uniqueValues[i][0];
  //     });

  //     //Create a sheet for data 
  //     const newSheet = createSheet(sheetName+'_'+getColumnName(sheet, columnIndex)+'='+uniqueValues[i][0]);
  //     newSheet.getRange(HEADERS_ROW_INDEX, 1, header.length, header[0].length).setValues(header);
  //     newSheet.getRange(HEADERS_ROW_INDEX + 1, 1, filteredData.length, filteredData[0].length).setValues(filteredData);
  //     SpreadsheetApp.flush();
      
  //     //Get the first column for which we have data in this Part
  //     const first_data_col = get_first_column_part_data(header, uniqueValues[i][0], scanning_start_col);

  //     // BEGIN FIRST PASS
  //     // If there are empty data columns before this column, delete them
  //     // Empty columns determined by looking at first column after our "scanning start" var and our "first data col" var
  //      if (first_data_col - (scanning_start_col + 1) != 0) { // If there are empty columns
  //        // create variables for Google column numbers for bulk deletion because deleteColumns is 1-indexed, but apps script is 0-indexed
  //        // deleteColumns needs where to start deleting and how many columns to delete after
  //         var googleScanningStart = scanning_start_col + 2;
  //         var googleFirstDataCol = first_data_col + 1;
  //         var googleNumberColsToDelete = googleFirstDataCol - googleScanningStart;
  //       newSheet.deleteColumns(googleScanningStart, googleNumberColsToDelete); 
  //     }

  //     // Alert so we know this has been done
  //     Logger.log("I am 50% done working on " + newSheet.getSheetName());

  //     // BEGIN SECOND PASS
  //     // Because we've had some deletions, we need to update the header so we're indexing appropriately
  //     const second_pass_data = newSheet.getDataRange().getValues();
  //     const second_pass_header = [second_pass_data[0]];

  //     //Now get where we need to start deleting from and how many cols to delete in the second pass
  //     const second_pass_delete_start = get_first_column_without_data(second_pass_header, uniqueValues[i][0],  scanning_start_col);
  //     var lastColumn = newSheet.getLastColumn(); 
      
  //     // now bulk delete columns on the second pass
  //     if (lastColumn - second_pass_delete_start != 0) {
  //       // create variables for Google column numbers for bulk deletion because deleteColumns is 1-indexed, but apps script is 0-indexed
  //       var googleSecondPassDeleteStart = second_pass_delete_start + 1;
  //       var googleLastCol = lastColumn + 1;
  //       var googleSecondPassNumberColsDelete = (googleLastCol - googleSecondPassDeleteStart);
  //       newSheet.deleteColumns(googleSecondPassDeleteStart, googleSecondPassNumberColsDelete);
  //     }
  //     Logger.log("Yay! I am done working on " + newSheet.getSheetName());
  //  }
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
* Helper function for getting the name of the column.
*/
function getColumnName(sheet, column) {
 return sheet.getRange(HEADERS_ROW_INDEX, column).getValue();
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
