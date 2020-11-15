// OPEN CLOSE
// v2.21 Change order of sheets - move Notices to the back, move Rosters to the front
// v1.2


function openOrdering(msg){
  refreshFormulae()
  formatOrders()
  unlockSheet("Orders")
  
  var defaultMsg = "Open until " + CLOSE_DAY + " at " + CLOSE_TIME
  setStatus(msg || defaultMsg)
  hideAdminSheets()
  
  // notify members?
}


function closeOrdering(msg){
  lockSheetByName("Orders", "Ordering Closed")
  setStatus(msg || "Ordering closed")
  deleteTriggers()
}


function setStatus(msg){
  SpreadsheetApp.getActiveSpreadsheet().getRangeByName("ord_Status").setValue(msg)
}

function deleteTriggers(){// deletes ALL triggers on this spreadsheet belonging to this user
  ss = SpreadsheetApp.getActiveSpreadsheet()
  var triggers = ScriptApp.getUserTriggers(ss)
  
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
}

function unlockSheet(name){ // Remove sheet protection
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  var protection = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET)[0];
  if (protection && protection.canEdit()) {
    protection.remove();
  }
}

function lockSheetByName(name, desc){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name)
  lockSheet(sheet, desc)
}


function lockSheet(sheet, desc){ // Add sheet protection
 var protection = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
 if (protection[0] !== undefined && protection[0].canEdit()) {
   for (var i =0; i < protection.length; i++){
     Logger.log(protection[i].getDescription())
     if (protection[i].getDescription() == 'Edit when locked'){// copy protection
       var editors = protection[i].getEditors()
       var sProtection = sheet.protect().setDescription(desc)
           
       // Ensure the current user is an editor before removing others. Otherwise, if the user's edit
       // permission comes from a group, the script will throw an exception upon removing the group.
       var me = Session.getEffectiveUser();
       sProtection.addEditor(me);
       sProtection.removeEditors(sProtection.getEditors());  // removes everyone except this user and owner
       sProtection.addEditors(editors)
     }
   }
 } else {
   // lock it
   sheet.protect().setDescription(desc)
 }
}

function setAll(arr, v) {// copy v to whole single dimension array
    var i, n = arr.length;
    for (i = 0; i < n; ++i) {
        arr[i] = v;
    }
}

function formatOrders() {// hide column F and colour the columns
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Orders")
  
  // hide column F 
  if (isFresh()) {
    sheet.hideColumns(6)
  }
  
  var range = SpreadsheetApp.getActiveSpreadsheet().getRangeByName("ord_Data")

  var firstRow = range.getRow()
  var numRows = range.getNumRows()
  var numCols = range.getNumColumns()
  var colours = []
  
  var col1 = new Array(numRows)
  setAll(col1, "#E7E7F0")
  
  var col2 = new Array(numRows)
  setAll(col2, "antiquewhite" )
  
  var col3 = new Array(numRows)
  setAll(col3, "white")
  
  var col4 = new Array(numRows)
  setAll(col4, "beige")
  
  for (var i = 0; i < numCols; i++){
    switch(i % 4) {
      case 0: colours[i] = col1; break;
      case 1: colours[i] = col3; break;
      case 2: colours[i] = col2; break;
      case 3: colours[i] = col4
    }
  
  }
   // set the last row to darkcyan 
  var  tColours  = ArrayLib.transpose(colours)
  
  var row = new Array(numCols)
  setAll(row, "#003333")
  tColours[numRows-1] = row
  range.setBackgrounds(tColours)
  


 
  
  
}

function tidyUpSheets() {
  hideAdminSheets()
  sortSheets()
}

function sortSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var ordered = ["Roster", "Orders",  "Totals", "Banking", "Members", "FAQ", "Notices"]   //"Pre-tweak Orders" , "Bank Acct Details", "FreshDirect Order"] 
  
  for (var i = 0; i < ordered.length; i++){
    try {
      ss.getSheetByName(ordered[i]).activate()
      ss.moveActiveSheet(i+1)       // sheets are indexed from 1
    } 
    catch (error) {// if sheet ain't there we don't care }
    }
  }

  ss.getSheetByName("Notices").activate()
}


function hideAdminSheets(){
  var visibles = ['Notices', 'Orders','Pre-tweak Orders', 'Totals', 'Banking', 'Payers', 'Roster', 'Members', 'FAQ']
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets()
  for (var i =0; i<sheets.length; i++){
    if (visibles.indexOf(sheets[i].getName()) == -1){sheets[i].hideSheet()}
  }
}
