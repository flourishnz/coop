

function ZeroThisRow(){
  var range = getRangeOrdersThisRow()
  var data = range.getValues();
  
  for (v = 0; v < data[0].length; v++ ){
   if (data[0][v]) {                    // for each value (order) found (no matter what value it is)
     data[0].splice(v,1,0)              // starting at v, replace 1 value with the value 0
   } 
 }
                
  range.setValues(data);                  // copy the changed row back to the sheet
}


function go(){
  //addSpecialToSheet(specialProduct)
  getRangeOrdersThisProduct("Bananas", "Pre-tweaked Orders")
}

function addSpecialToSheet(product){// incomplete
  var ss = SpreadsheetApp.openById(targetSSID);
  var specialSheet = ss.getSheetByName(product) 
  if (specialSheet == null) return;
  
  var orders = ss.getSheetByName('Orders')
  var pretweaks = ss.getSheetByName('Pre-tweaked Orders')
  var data = specialSheet.getRange(2, 2, specialSheet.getLastRow()-1, 2).getValues()
  
  for (var i = 0; i < data.length; i++){
    
  }
}

function getRangeMembersFromOrders() {//...still need to test this one
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Orders");                                   
  var numberofcolumns = sheet.getLastColumn()- FIRST_ORDER_COLUMN + 1;  // and continue to the end of the ro  
  return range = sheet.getRange(thisrow, FIRST_ORDER_COLUMN, 1, numberofcolumns);
}


function getRangeOrdersThisRow() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var thisrow = sheet.getActiveCell().getRow();
  var numberofcolumns = sheet.getLastColumn()- FIRST_ORDER_COLUMN + 1;  // and continue to the end of the row  
  return range = sheet.getRange(thisrow, FIRST_ORDER_COLUMN, 1, numberofcolumns);
}


function getRangeOrdersThisProduct(product ,optSheetName) {// writing this one
  var sheetName = optSheetName || "Orders"
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName)
  var data = sheet.getDataRange().getValues()
  //var index = ArrayLib
}