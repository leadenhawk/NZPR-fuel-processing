function onOpen() {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
    
  var filterSheetMenu = [];
  filterSheetMenu = [
    {name: "Filter (show/hide)", functionName: "filter"},
    {name: "Filter (do not show/hide)", functionName: "filterNSH"},
    {name: "Show All", functionName: "showAll"},
    {name: "Sort by Card",functionName: "sortByCard"},
    {name: "Sort by Client", functionName: "sortByClient"},
    {name: "Sort by Card, then Client", functionName: "sortByCardThenClient" },
    {name: "Move", functionName: "moveFilteredLine" },
    {name: "export as csv files", functionName: "saveAsCSV"}
  ];
  
  ss.addMenu("Filter", filterSheetMenu);
  
}


// variables for the filtering function
var targetSheet =  "Sheet1";      //"testing";                                     // change the target sheet
var csvSheet = "Proc";                                                             // change the destination sheet
var theClientColumn = 5;//0;                                                           // change client column
var theCardColumn = 4;//2;                                                             // change card column
var theSelectedColumn = 8;//3;                                                         // change selected column

function filter() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(targetSheet);
  var data = sheet.getDataRange().getValues();

  var clientColumn = theClientColumn;//0;
  var cardColumn = theCardColumn;//2;
  var selectedColumn = theSelectedColumn;//3; 
  
 
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
  //sortBySelected();
  goToTop();
}








function filterNSH() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(targetSheet);
  var data = sheet.getDataRange().getValues();

  var clientColumn = theClientColumn;//0;
  var cardColumn = theCardColumn;//2;
  var selectedColumn = theSelectedColumn;//3; 
  
 
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
    //sheet.hideRows(i+1);
    var cell = sheet.getRange(i+1,selectedColumn+1);
    cell.setValue("");
  }
  
    // FILTERING SECTION OF CODE
    //iterate over all rows
    for(var i=3; i< data.length; i++){
            
      if  ((col1.length > 0)  &&  (col2.length == 0)) {
        if ( col1.includes(data[i][clientColumn])) {
          //sheet.showRows(i+1);
          
          var cell = sheet.getRange(i+1,selectedColumn+1);
          cell.setValue("X");
        }    
      } 
      
      else if((col1.length == 0) && ( col2.length > 0)) {
        if ( col2.includes(data[i][cardColumn])) {
          //sheet.showRows(i+1);
          
          var cell = sheet.getRange(i+1,selectedColumn+1);
          cell.setValue("X");
        }
      }
      
      else if(( col1.length > 0) && ( col2.length > 0)) {
        if (( col1.includes(data[i][clientColumn]) ) && ( col2.includes(data[i][cardColumn]) )){
          //sheet.showRows(i+1);
          
          var cell = sheet.getRange(i+1,selectedColumn+1);
          cell.setValue("X");
        }
      }
      
      else if ((col1.length == 0) && ( col2.length == 0)) {
        //sheet.showRows(i+1);
        
        var cell = sheet.getRange(i+1,selectedColumn+1);
        cell.setValue("X");
      }
    }
  sortBySelected();
  goToTop();
}






function showAll() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(targetSheet);
  
  var maxRows = sheet.getMaxRows();

  //clear all the selection markers 
  var selectedColumn = theSelectedColumn;//3;
  
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
  var sheet = ss.getSheetByName(targetSheet);
  sheet.getRange('A4').activate();  
}


function sortByCardThenClient(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(targetSheet);
  
  var data = sheet.getDataRange().getValues();
  
  var range = sheet.getRange("A4:C500");
  range.sort([{column: theCardColumn+1, ascending: true}, {column: theClientColumn+1, ascending: true}]);
}

function sortByCard(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(targetSheet);
  var data = sheet.getDataRange().getValues();
  
  var rows = sheet.getDataRange();
  var numRows = rows.getNumRows() + 10;
  
  var range = sheet.getRange("A4:C"+numRows);
  Browser.msgBox('under development', 'the range is A4:C'+numRows, Browser.Buttons.OK);
  //range.sort(theCardColumn+1);
}

function sortByClient(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(targetSheet);
  var data = sheet.getDataRange().getValues();
  var range = sheet.getRange("A4:C500");
  range.sort(theClientColumn+1);
}

function sortBySelected(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(targetSheet);
  var data = sheet.getDataRange().getValues();
  var range = sheet.getRange("A4:C500");
  range.sort(theSelectedColumn+1);
}

function moveFilteredLine(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var firstSheet = ss.getSheetByName(targetSheet);
  var secondSheet = ss.getSheetByName(csvSheet);
  
  var selectedColumn = theSelectedColumn;//3;
  
  var data = firstSheet.getDataRange().getValues();
  var numberOffirstSheetRows = firstSheet.getLastRow();
  for (var i = 3; i < numberOffirstSheetRows; i++) { 
    if (data[i][selectedColumn] != ""){//checks if the row is marked/checked/approved
      Logger.log(data[i][selectedColumn])
      secondSheet.appendRow([data[i][2],       // date                                                    // change what is moved & the order
                             data[i][0],       // vendor
                             data[i][1],       // fuel_type
                             data[i][3],       // amount
                             data[i][5]        // client
                            ]);
      
      clearFirstSheet(); //remove copied line from sheet
    }
  }
  goToTop();
}

function clearFirstSheet() {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(targetSheet);
  
  var selectedColumn = theSelectedColumn;//3;
  
  var rows = sheet.getDataRange();
  var numRows = rows.getNumRows();
  var values = rows.getValues();
  
  var data = sheet.getDataRange().getValues();
  
  var rowsDeleted = 0;
  for (var i = 3; i <= numRows - 1; i++) {//i=3 so the headers are not deleted
    var row = values[i];
    if (row[theSelectedColumn] != '') {//checks to make sure only rows with something in the "selected" column are selected 
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
