HEADERS_ROW_INDEX = 1
P_COL_START = 27

function start() {
  // Modify the three lines below for your own data sheet
  var sheetName = 'Sheet1'; // Name of the sheet that you imported //UPDATE THIS TO 
  const columnName = 'Part'; // This should be Part usually
  const columnsToExclude = ['Education Level', 'Country']; // List all column names that you would like to remove
////DO NOT MODIFY ANYTHING BELOW unless you know what you are doing :D ////////////////////

  P_COL_START -= columnsToExclude.length;

  var sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
  var tempOriginalSheet = sheet.copyTo(SpreadsheetApp.getActiveSpreadsheet());
  tempOriginalSheet.setName("processed_" + sheetName);

  sheetName = "processed_" + sheetName;
  sheet = tempOriginalSheet;


  const columnsToExcludeInIndex = []
  const data = sheet.getDataRange().getValues();
  const header = [data[0]];
  for (var i = 0; i < columnsToExclude.length; i++) {
    columnsToExcludeInIndex.push(find_column_index_from_column_name(header[0], columnsToExclude[i]));
  }
  columnsToExcludeInIndex.sort();
  Logger.log(columnsToExcludeInIndex);
  for (var j = columnsToExcludeInIndex.length - 1; j >=0; j--) {
    sheet.deleteColumn(columnsToExcludeInIndex[j]);
  }

  run_group_by(sheetName, columnName);
}

function find_column_index_from_column_name(header, columnName) {
  for (var i = 0; i < header.length; i++) {
    if (header[i] == columnName) {
      return i+1;
    }
  }
}

function run_group_by(sheetName, columnName) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
  const data = sheet.getDataRange().getValues();
  const header = [data[0]];

  const columnIndex = find_column_index_from_column_name(header[0], columnName)

  const uniqueValues = getUniqueColumn(sheet, columnIndex);
  Logger.log("Extracted header: " + header);

  for (var i = 0; i < uniqueValues.length; i++) {
    const filteredData = data.filter(function(row) {
        return row[columnIndex-1] == uniqueValues[i][0];
    });
    const newSheet = createSheet(sheetName+'_'+getColumnName(sheet, columnIndex)+'='+uniqueValues[i][0]);
    newSheet.getRange(HEADERS_ROW_INDEX, 1, header.length, header[0].length).setValues(header);
    newSheet.getRange(HEADERS_ROW_INDEX + 1, 1, filteredData.length, filteredData[0].length).setValues(filteredData);
    SpreadsheetApp.flush();
    const columns_to_delete = extract_columns_to_delete(header, uniqueValues[i][0]);
    Logger.log(columns_to_delete);
    for (var j = columns_to_delete.length - 1; j >=0; j--) {
      newSheet.deleteColumn(columns_to_delete[j]+1);
    }
    Logger.log("# " + filteredData.length + " rows are written to the sheet: " + newSheet.getSheetName());
  }
}

function extract_columns_to_delete(header, part_num) {
  var columns = [];
  for (var i = P_COL_START-1; i < header[0].length; i++) {
    if (header[0][i].indexOf("P"+part_num+":") == -1) {
      columns.push(i);
    }
  }
  return columns;
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

    Logger.log(yourNewSheet.getSheetName() + " - new sheet is created!");

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
const getUnique_ = array2d => [...new Set(array2d.flat())];

/*
 * Get a list of unique values from a given column.
 */
function getUniqueColumn(sheet, column) {
  const rg = sheet.getRange(1 + HEADERS_ROW_INDEX, column, sheet.getLastRow() - HEADERS_ROW_INDEX, 1);
  const uniqueValues = getUnique_(rg.getValues()).map(e => [e]);
  Logger.log("Unique Values for the header: '" + getColumnName(sheet, column) + "' => " + uniqueValues);
  return uniqueValues;
}