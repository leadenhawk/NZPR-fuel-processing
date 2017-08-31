function onOpen() {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
    
  var filterSheetMenu = [];
  filterSheetMenu = [
    {name: "Filter", functionName: "filter"},
    {name: "Show All", functionName: "showAll"},
    {name: "Sort by Card",functionName: "sortByCard"},
    {name: "Sort by Client", functionName: "sortByClient"},
    {name: "Sort", functionName: "sortByCardThenClient" },
    {name: "Move", functionName: "moveFilteredLine" },
    {name: "export as csv files", functionName: "saveAsCSV"}
  ];
  
  ss.addMenu("Filter", filterSheetMenu);
  
}

// this is to make the "includes" method work!! 
// polyfill from https://developer.mozilla.org/en/docs/Web/JavaScript/Reference/Global_Objects/Array/includes
// https://tc39.github.io/ecma262/#sec-array.prototype.includes
if (!Array.prototype.includes) {
  Object.defineProperty(Array.prototype, 'includes', {
    value: function(searchElement, fromIndex) {

      // 1. Let O be ? ToObject(this value).
      if (this == null) {
        throw new TypeError('"this" is null or not defined');
      }

      var o = Object(this);

      // 2. Let len be ? ToLength(? Get(O, "length")).
      var len = o.length >>> 0;

      // 3. If len is 0, return false.
      if (len === 0) {
        return false;
      }

      // 4. Let n be ? ToInteger(fromIndex).
      //    (If fromIndex is undefined, this step produces the value 0.)
      var n = fromIndex | 0;

      // 5. If n ≥ 0, then
      //  a. Let k be n.
      // 6. Else n < 0,
      //  a. Let k be len + n.
      //  b. If k < 0, let k be 0.
      var k = Math.max(n >= 0 ? n : len - Math.abs(n), 0);

      function sameValueZero(x, y) {
        return x === y || (typeof x === 'number' && typeof y === 'number' && isNaN(x) && isNaN(y));
      }

      // 7. Repeat, while k < len
      while (k < len) {
        // a. Let elementK be the result of ? Get(O, ! ToString(k)).
        // b. If SameValueZero(searchElement, elementK) is true, return true.
        // c. Increase k by 1. 
        if (sameValueZero(o[k], searchElement)) {
          return true;
        }
        k++;
      }

      // 8. Return false
      return false;
    }
  });
}


function filter() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("testing");
  
  var data = sheet.getDataRange().getValues();
  
  var clientColumn = 0;
  var cardColumn = 2;
  var selectedColumn = 3;
  
  var empty = (""||0||undefined);

  
  
  // getting the filters in arrays
  var col1 = [];
  col1 = sheet.getRange('B1:1').getValues().join().split(',').filter(Boolean);
  
  var col2 = [];
  col2 = sheet.getRange('B2:2').getValues().join().split(',').filter(Boolean);
  
  
  // card numbers must be strings for "contains" method to work (!)
  for (var i=0; i < col2.length; i++){
    col2[i]=col2[i].toString();
  }
  
  for(var i=3; i< data.length; i++){
     data[i][cardColumn] = data[i][cardColumn].toString();
  }
    
  // hide all rows and remove all selectors before running filtering section of code
  for(var i=3; i< data.length; i++){
    sheet.hideRows(i+1);
    var cell = sheet.getRange(i+1,selectedColumn+1);
    cell.setValue("");
  }
  
    // FILTERING SECTION OF CODE
    //iterate over all rows
    for(var i=3; i< data.length; i++){
            
      if  ((col1.length > 0)  &&  (col2.length == 0)) {
        if ( col1.includes(data[i][clientColumn])) {
          sheet.showRows(i+1);
          
          var cell = sheet.getRange(i+1,selectedColumn+1);
          cell.setValue("X");
        }    
      } 
      
      else if((col1.length == 0) && ( col2.length > 0)) {
        if ( col2.includes(data[i][cardColumn])) {
          sheet.showRows(i+1);
          
          var cell = sheet.getRange(i+1,selectedColumn+1);
          cell.setValue("X");
        }
      }
      
      else if(( col1.length > 0) && ( col2.length > 0)) {
        if (( col1.includes(data[i][clientColumn]) ) && ( col2.includes(data[i][cardColumn]) )){
          sheet.showRows(i+1);
          
          var cell = sheet.getRange(i+1,selectedColumn+1);
          cell.setValue("X");
        }
      }
      
      else if ((col1.length == 0) && ( col2.length == 0)) {
        sheet.showRows(i+1);
        
        var cell = sheet.getRange(i+1,selectedColumn+1);
        cell.setValue("X");
      }
    }
  goToTop();
}


function showAll() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("testing");
  
  var maxRows = sheet.getMaxRows();

  //clear all the selection markers 
  var selectedColumn = 3;
  
  for(var i=3; i< maxRows; i++){
    var cell = sheet.getRange(i+1,selectedColumn+1);
    cell.setValue("");
  }
  
  //show all the rows
  sheet.showRows(1, maxRows);
  
  goToTop();
}



function goToTop(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("testing");
  sheet.getRange('A4').activate();  
}


function sortByCardThenClient(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("testing");
  
  var data = sheet.getDataRange().getValues();
  
  var range = sheet.getRange("A4:C500");
  range.sort([{column: 3, ascending: true}, {column: 2, ascending: true}]);

}

function sortByCard(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("testing");
  var data = sheet.getDataRange().getValues();
  var range = sheet.getRange("A4:C500");
  range.sort(3);
}

function sortByClient(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("testing");
  var data = sheet.getDataRange().getValues();
  var range = sheet.getRange("A4:C500");
  range.sort(2);
}


function moveFilteredLine(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var firstSheet = ss.getSheetByName("testing");
  var secondSheet = ss.getSheetByName("Proc");
  
  var selectedColumn = 3;
  
  var data = firstSheet.getDataRange().getValues();
  var numberOffirstSheetRows = firstSheet.getLastRow();
  for (var i = 3; i < numberOffirstSheetRows; i++) { 
    if (data[i][selectedColumn] != ""){//checks if the row is marked/checked/approved
      Logger.log(data[i][selectedColumn])
      secondSheet.appendRow([data[i][0],            
                             data[i][1],
                             data[i][2]
                            ]);
      
      clearFirstSheet(); //remove copied line from sheet
    }
  }
  goToTop();
}

function clearFirstSheet() {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("testing");
  
  var selectedColumn = 3;
  
  var rows = sheet.getDataRange();
  var numRows = rows.getNumRows();
  var values = rows.getValues();
  
  var data = sheet.getDataRange().getValues();
  
  var rowsDeleted = 0;
  for (var i = 3; i <= numRows - 1; i++) {//i=3 so the headers are not deleted
    var row = values[i];
    if (row[3] != '') {//checks to make sure only rows with something in the "selected" column are selected
      sheet.deleteRow((parseInt(i)+1) - rowsDeleted);
      rowsDeleted++;
    }
  }
}


/*
 * script to export data in all sheets in the current spreadsheet as individual csv files
 * files will be named according to the name of the sheet
 * author: Michael Derazon
*/

//function onOpen() {
//  var ss = SpreadsheetApp.getActiveSpreadsheet();
//  var csvMenuEntries = [{name: "export as csv files", functionName: "saveAsCSV"}];
//  ss.addMenu("csv", csvMenuEntries);
//};

function saveAsCSV() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  // create a folder from the name of the spreadsheet
  var folder = DriveApp.createFolder(ss.getName().toLowerCase().replace(/ /g,'_') + '_csv_' + new Date().getTime());
  for (var i = 0 ; i < sheets.length ; i++) {
    var sheet = sheets[i];
    // append ".csv" extension to the sheet name
    fileName = sheet.getName() + ".csv";
    // convert all available sheet data to csv format
    var csvFile = convertRangeToCsvFile_(fileName, sheet);
    // create a file in the Docs List with the given name and the csv data
    folder.createFile(fileName, csvFile);
  }
  Browser.msgBox('Files are waiting in a folder named ' + folder.getName());
}

function convertRangeToCsvFile_(csvFileName, sheet) {
  // get available data range in the spreadsheet
  var activeRange = sheet.getDataRange();
  try {
    var data = activeRange.getValues();
    var csvFile = undefined;

    // loop through the data in the range and build a string with the csv data
    if (data.length > 1) {
      var csv = "";
      for (var row = 0; row < data.length; row++) {
        for (var col = 0; col < data[row].length; col++) {
          if (data[row][col].toString().indexOf(",") != -1) {
            data[row][col] = "\"" + data[row][col] + "\"";
          }
        }

        // join each row's columns
        // add a carriage return to end of each row, except for the last one
        if (row < data.length-1) {
          csv += data[row].join(",") + "\r\n";
        }
        else {
          csv += data[row];
        }
      }
      csvFile = csv;
    }
    return csvFile;
  }
  catch(err) {
    Logger.log(err);
    Browser.msgBox(err);
  }
}

