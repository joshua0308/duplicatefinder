// this is a Google Apps Script project

function onOpen() {
  var spreadsheet = SpreadsheetApp.getActive();
  var menuItems = [
    { name: 'Find duplicates...', functionName: 'findDuplicate' },
    { name: 'Remove duplicates...', functionName: 'removeDuplicate' }
  ];
  spreadsheet.addMenu('Duplicates', menuItems);
}

function removeDuplicate() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = SpreadsheetApp.getActiveSheet();
  var range = sheet.getActiveRange();
  var data = range.getValues();

  var rowNum = range.getRow();
  var columnNum = range.getColumn();
  var columnLength = data[0].length;

  var uniqueData = [];
  var duplicateData = [];

  // iterate through each 'row' of the selected range
  // x is
  // y is
  var x = 0;
  var y = data.length;

  // when row is
  while (x < y) {
    var row = data[x];
    var duplicate = false;

    // iterate through the uniqueData array to see if 'row' already exists
    for (var j = 0; j < uniqueData.length; j++) {
      if (row.join() == uniqueData[j].join()) {
        // if there is a duplicate, delete the 'row' from the sheet and add it to the duplicateData array
        duplicate = true;
        var duplicateRange = sheet.getRange(
          rowNum + x,
          columnNum,
          1,
          columnLength
        );
        duplicateRange.deleteCells(SpreadsheetApp.Dimension.ROWS);
        duplicateData.push(row);

        // rows shift up by one when duplicate is deleted
        // in effect, it skips a line
        // so we need to decrement x to stay in the same line
        x--;
        y--;
        range = sheet.getActiveRange();
        data = range.getValues();
        // return;
      }
    }

    // if there are no duplicates, add 'row' to the uniqueData array
    if (!duplicate) {
      uniqueData.push(row);
    }
    x++;
  }

  // create a new sheet with the duplicate data
  if (duplicateData) {
    var newSheet = spreadsheet.insertSheet();
    var header = duplicateData.length + ' duplicates found';
    newSheet.setName('Duplicates v' + (spreadsheet.getSheets().length - 1));
    newSheet.appendRow([header]);

    for (var k = 0; k < duplicateData.length; k++) {
      newSheet.appendRow(duplicateData[k]);
    }
  }
}

function findDuplicate() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var range = sheet.getActiveRange();
  var data = range.getValues();

  var rowNum = range.getRow();
  var columnNum = range.getColumn();
  var columnLength = data[0].length;

  var uniqueData = [];

  // iterate through each 'row' of the selected range
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var duplicate = false;

    // iterate through the uniqueData array to see if 'row' already exists
    for (var j = 0; j < uniqueData.length; j++) {
      if (row.join() == uniqueData[j].join()) {
        // if there is a duplicate, highlight the 'row' from the sheet
        duplicate = true;
        var duplicateRange = sheet.getRange(
          rowNum + i,
          columnNum,
          1,
          columnLength
        );
        duplicateRange.setBackground('yellow');
      }
    }

    // if there are no duplicates, add 'row' to the uniqueData array
    if (!duplicate) {
      uniqueData.push(row);
    }
  }
}
