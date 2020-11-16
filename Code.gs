// CODE.GS

// v1.2 add MEMBERSHIP_BOND constant
// v1.1  Move email addresses to globals
// v1.07 Add "All pack day reports" option to menu
// v1.06 Adjust for new Dry Members sheet layout
// v1.05 "Add Members" option added to menu, removed some commented code
// v1.04 Moved sharePdfPacksheets here for now because it is called by several reports
// v1.03

// 26 Aug 18 Added stocktake to Dry menu
// 15-Jun-18 added functions say() and rounded(); relocated getLatestSS, getLatestSsId etc - not in use??
//  1-Mar-18 add MIN_ORDER_FEE constant
// 11-Feb-18 Synched All - adding reports to Dry Menu - Adding constants
// 28-Nov-17 SYNCHED Fresh, A and B
// added mem_id_offset constant
// 29-Sep-17 SYNCHED Fresh and DRY B
//
//           added membership_fee constant
//  1-Feb-17 use constants
// 30-Jul-16 menu edited
// 25-Jul-16 synched
// 17-Jul-16 added onEdit - Change Log
//--------------------------------------

const _ = LodashGS.load()


if (isFresh()) { 
  isFRESH = true
  isDRY = false
  VENDOR_COLUMN = 2
  PRODUCT_COLUMN = 3
  UNIT_COLUMN = 4
  CRATE_COUNT_COLUMN = 5
  PRICE_EXCL_COLUMN = 6
  PRICE_COLUMN = 7
  TOTAL_KGS_COLUMN = 8
  TOTAL_CRATES_COLUMN = 9

  FIRST_ORDER_COLUMN = 10
  FIRST_ORDER_ROW = 5
  
  USERNAME_ROW = 3
  USERID_ROW = 4

  MEMBERSHIP_FEE = 75
  MEMBERSHIP_BOND = 50
  MEM_ID_OFFSET = 1 // MEMBERS SHEET
  MEM_MOBILE_OFFSET = 5  // MEMBERS SHEET
  TOT_ID_ROW = 3  // TOTALS SHEET
  
  CLOSE_DAY = "Monday"
  CLOSE_TIME = "9:00 pm"
  MIN_ORDER_FEE = 3

  COOP_EMAIL = "kapitifresh.co.op@gmail.com"
  
  LOCAL_EMAIL = "matt.mcrae86@gmail.com"
  LOCAL_NAME = "Tiffany"
  
  PRICES_EMAIL = "mattrobin24@gmail.com"
  PRICES_NAME = "Matt"
  
  TREASURER_EMAIL = "4fruitandveg@gmail.com"
  TREASURER_NAME = "Joanne"
  
  ROSTERS_EMAIL = "info@carolshortis.com"
  ROSTERS_NAME = "Carol"
  
  MEMBERSHIP_EMAIL = "mooksjoinza@gmail.com"
  MEMBERSHIP_NAME = "Amanda"
  
  IT_EMAIL = "flourish.nz@gmail.com"
  IT_NAME = "Julie"
}

if (isDry()) {
  isFRESH = false
  isDRY = true
  FIRST_ORDER_COLUMN = 9
  FIRST_ORDER_ROW = 6
  
  PRICE_COLUMN = 7
  PRODUCT_COLUMN = 2
  UNIT_COLUMN = 3
  GROUP_COLUMN = 1
  
  USERID_ROW = 4
  USERNAME_ROW = 3
  
  MEMBERSHIP_FEE = 25
  MEMBERSHIP_BOND = 0
  MEM_ID_OFFSET = 0 // MEMBERS SHEET
  MEM_MOBILE_OFFSET = 3  // MEMBERS SHEET
  TOT_ID_ROW = 2  // TOTALS SHEET

  CLOSE_DAY = "Monday"
  CLOSE_TIME = "6:00 pm"
  MIN_ORDER_FEE = 2
  
  COOP_EMAIL = "affordableorganics07@gmail.com"
  
  TREASURER_EMAIL = "affordableorganics07@gmail.com"
  ROSTERS_EMAIL = "affordableorganics07@gmail.com"
  MEMBERSHIP_EMAIL = "affordableorganics07@gmail.com"
  
  IT_EMAIL = "kapitidry.coop@gmail.com"
  IT_NAME = "Julie"

//  VENDOR_COLUMN is undefined
}

brbr = "<br><br>"


function onOpen() {

  if (isFRESH){
    SpreadsheetApp.getUi()
    
    .createMenu('Co-op Admin')
//      .addItem('Open Ordering', 'openOrdering')
//      .addItem('Send Reminders', 'sendReminderSMS')')
//      .addItem('Close Ordering', 'closeOrdering')
      .addItem('Notify this member of balance', 'notifyThisMemberOfBalance')

    
      .addSubMenu(SpreadsheetApp.getUi()
                  .createMenu('Reports')
                    .addItem('Both pack reports', 'runFreshReports')
                    .addItem('Pack list', 'createReportFreshPacklist')
                    .addItem('Bin list', 'createReportBinList')
               )
        
      .addSubMenu(SpreadsheetApp.getUi()
                  .createMenu('Structural')
                    .addItem('Add members', 'addMembers')
                    .addItem('Remove this member', 'removeThisMember')
                    .addItem('Rollover', 'rollover')
                    .addItem('Refresh Formulae', 'refreshFormulae')
                    .addItem('Test statements', 'testStatements')
                 )
      .addToUi();
    
    SpreadsheetApp.getUi()
    .createMenu('Tweaking')
      .addItem('Close Ordering and Start Tweaking', 'startTweaking')     
      .addItem('Zero out these products', 'zeroOutSelectedRows')
      .addItem('Reinstate this product', 'reinstateRow')
      .addItem('Summarise this product', 'summariseThis')
      .addItem('Finish Tweaking', 'doneTweaking')   
    .addToUi();
  }
  else {
    SpreadsheetApp.getUi()
    
    .createMenu('Co-op Admin')
      .addItem('Open Ordering', 'openOrdering')
      .addItem('Send Reminders', 'sendReminderSMS')
      .addItem('Close Ordering', 'closeOrdering')                                      
      .addItem('Zero out this product', 'zeroOutSelectedRows')
      .addItem('Reinstate this product', 'reinstateRow')
    
    .addSubMenu(SpreadsheetApp.getUi()
                .createMenu('Reports')
                .addItem('All pack reports', 'runDryReports')
                .addItem('Member orders', 'createReportDryPickupLists')
                .addItem('Pickup checklist', 'createReportPickupChecklist')
                .addItem('Stocktake list', 'createReportStocktake')
               )
    
    .addSubMenu(SpreadsheetApp.getUi()
                .createMenu('Structural')
                .addItem('Add members', 'addMembers')
                .addItem('Remove this member', 'removeThisMember')
                .addItem('Rollover', 'rollover')
                .addItem('Refresh Formulae', 'refreshFormulae')
                .addItem('Tidy Up', 'tidyUpSheets')
                .addItem('Temp','temp')
               )
    .addToUi();
    
  }
}


 
function onEdit(e){
  var srcSheet = e.range.getSheet()
  if (srcSheet.getName() === "Orders"){
    var logSheet = e.source.getSheetByName("Change Log")
    var col = e.range.getColumn()
    var row = e.range.getRow()
    var editedUserId = srcSheet.getRange(USERID_ROW, col).getValue()
    var editedUserName = srcSheet.getRange(USERNAME_ROW, col).getValue()
    var product = srcSheet.getRange(row, PRODUCT_COLUMN).getValue()

    //var newValue = e.value
//    if (isNumeric(e.value)){// could be object representing previous value
//      var newValue = e.value} 
//    else {
//      var newValue = ""
//    }
    var newValue = (typeof e.value == "object" ? e.range.getValue() : e.value)
    
    var oldValue = e.oldValue || ""                  // e.oldValue could be "undefined"
    var entry = [new Date(), 
                 e.range.getA1Notation(),
                 editedUserId, 
                 editedUserName,
                 product,
                 oldValue,
                 newValue
                ]
    logSheet.appendRow(entry)
    
    makeToast(e)
  } 
  else 
    if (srcSheet.getName() === "Members"){
      var newValue = (typeof e.value == "object" ? e.range.getValue() : e.value); 
      var oldValue = e.oldValue || ""                  // e.oldValue could be "undefined"

      var col = e.range.getColumn()
      var row = e.range.getRow()
      var editedId = srcSheet.getRange(row, MEM_ID_OFFSET+1).getValue()
      var editedName = srcSheet.getRange(row, MEM_ID_OFFSET+2).getValue()
//      log(["Member change detected", editedName, editedId, 'Old: ' + oldValue, 'New: ' + newValue ])

      if (isValidId(editedId) && editedName.length > 0) {
        var member = getMember(row)
        CoopCoopLib.addMemberToContacts(member)
        log(["Member updated", editedName, editedId, 'Old: ' + oldValue, 'New: ' + newValue])
      }
    }
}
  

function makeToast(e){
  if (isNumeric(e.value)){   
    var sheet = e.range.getSheet()
    var data = sheet.getRange(e.range.getRow(), 1, 1, FIRST_ORDER_COLUMN - 1).getValues()[0]
    var product = data[PRODUCT_COLUMN - 1]
    var unit = data[UNIT_COLUMN - 1]
    var price = data[PRICE_COLUMN - 1]
    var qty = e.value
    var total = qty * price 
    if (unit.match(/kg/i)){
      var msg =  qty + " kg at " + '$'  +  price.toFixed(2)  + " /kg :  $ " + total.toFixed(2)
    }
    else if (unit.match(/ea|ct/i)){
      var msg =  qty + " at " + '$'  +  price.toFixed(2)  + " ...Total :  $ " + total.toFixed(2)
    } 
    else {
      e.source.toast("no match")}
    msg && e.source.toast(msg, product , 12)
  }
}

function setChangeLog(e){
  var logSheet = e.source.getSheetByName("Change Log")
  var col = e.range.getColumn()
  var row = e.range.getRow()
  var editedUserId = srcSheet.getRange(USERID_ROW, col).getValue()
  var editedUserName = srcSheet.getRange(USERNAME_ROW, col).getValue()
  var product = srcSheet.getRange(row, PRODUCT_COLUMN).getValue()
  Logger.log("get here 1")
  var newValue = e.value
  if (isNumeric(e.value)){// could be object representing previous valu!
    var newValue = e.value} 
  else {
    var newValue = ""
    }
  
  var oldValue = e.oldValue || ""                  // e.oldValue could be "undefined"
  var entry = [new Date(), 
               e.range.getA1Notation(),
               editedUserId, 
               editedUserName,
               product,
               oldValue,
               newValue
              ]
  logSheet.appendRow(entry)

}

function zeroOutRow() {
  // set all the orders in the current row to 0, starting at FIRST_ORDER_COLUMN
  var sheet = SpreadsheetApp.getActiveSheet();
  var thisrow = sheet.getActiveCell().getRow();

  var numberofcolumns = sheet.getLastColumn()- FIRST_ORDER_COLUMN + 1;  // and continue to the end of the row  
  var range = sheet.getRange(thisrow, FIRST_ORDER_COLUMN, 1, numberofcolumns);
  var data = range.getValues();
  
 for(var v = 0; v < data[0].length; v++){
   if (data[0][v]) {                    // for each value (order) found (no matter what value it is)
     data[0].splice(v,1,0)              // starting at v, replace 1 value with the value 0
   } 
 }
                
  range.setValues(data);                  // copy the changed row back to the sheet
}


function zeroOutSelectedRows() {
  // set all the orders in the selected rows to 0, starting at FIRST_ORDER_COLUMN
  var sheet = SpreadsheetApp.getActiveSheet();
  var firstRow = sheet.getActiveRange().getRow();       //gets actual row number rather than row index in the range
  var numRows = sheet.getActiveRange().getNumRows();  
  var numColumns = sheet.getLastColumn()- FIRST_ORDER_COLUMN + 1;     // ...the columns with member data in them
  var range = sheet.getRange(firstRow, FIRST_ORDER_COLUMN, numRows, numColumns);
  var data = range.getValues();
    
  for (var r = 0; r < numRows; r++){
    for (c = 0; c < numColumns; c++){
      if (data[r][c]) {                    // for each value (order) found (no matter what value it is)
        data[r].splice(c,1,0)              // replace it with zero:  splice(c ,1 ,0) means starting at c, reomove 1 value and replace it with the value 0
      }
    }   
  } 
  range.setValues(data);                  // copy the changed row(s) back to the sheet
}




function reinstateRow() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var sheet = ss.getActiveSheet()
  var thisRow = sheet.getActiveCell().getRow()
  var product = sheet.getRange(thisRow, PRODUCT_COLUMN).getValue()
  var numberofcolumns = sheet.getLastColumn() - FIRST_ORDER_COLUMN
  
  var preTweaks = ss.getSheetByName("Pre-tweak Orders") 
  var tweakrow = (product !== "") && getProductRow(product, preTweaks) || !product && thisRow || 0
  
  if (!tweakrow) throw("Product missing from Pre-tweak Orders");
  
  var ptorder = preTweaks.getRange(tweakrow, FIRST_ORDER_COLUMN, 1, numberofcolumns);
  var order = sheet.getRange(thisRow, FIRST_ORDER_COLUMN, 1, numberofcolumns);
  order.setValues(ptorder.getValues()); 
}


//function reinstateRows() {//incomplete
//  get selected rows
//  get all pre-tweak rows
//  for each selected row
//    find pre-tweak row
//    replace data
//  
////  var ss = SpreadsheetApp.getActiveSpreadsheet()
////  var sheet = ss.getActiveSheet()
////  var thisRow = sheet.getActiveCell().getRow()
////  var product = sheet.getRange(thisRow, PRODUCT_COLUMN).getValue()
////  var numberofcolumns = sheet.getLastColumn() - FIRST_ORDER_COLUMN
////  
////  var preTweaks = ss.getSheetByName("Pre-tweak Orders") 
////  var tweakrow = (product !== "") && getProductRow(product, preTweaks) || !product && thisRow || 0
////  
//  var sheet = SpreadsheetApp.getActiveSheet();
//  var firstRow = sheet.getActiveRange().getRow();                     //gets actual row number rather than row index in the range
//  var numRows = sheet.getActiveRange().getNumRows();  
//  var numColumns = sheet.getLastColumn()- FIRST_ORDER_COLUMN + 1;     // ...the columns with member data in them
//  var range = sheet.getRange(firstRow, FIRST_ORDER_COLUMN, numRows, numColumns);
//  var data = range.getValues();
//  
//  var preTweaks = ss.getSheetByName("Pre-tweak Orders").getDataRange().getValues()
//  var product 
//  
//  
//  
////  if (!tweakrow) throw("Product missing from Pre-tweak Orders");
////  
////  var ptorder = preTweaks.getRange(tweakrow, FIRST_ORDER_COLUMN, 1, numberofcolumns);
////  var order = sheet.getRange(thisRow, FIRST_ORDER_COLUMN, 1, numberofcolumns);
////  order.setValues(ptorder.getValues()); 
//  
//  for (var r = 0; r < numRows; r++){
//    product = data[r, PRODUCT_COLUMN]                                //sheet.getRange(r, PRODUCT_COLUMN).getValue()
//    var tweakrow = (product !== "") && getProductRow(product, preTweaks) || !product && thisRow || 0
//      }
//  
//  } 
//  range.setValues(data);                  // copy the changed row(s) back to the sheet
//}



function testit(){
  Logger.log( getProductRow("Rhubarb"))

}

function getProductRow(product, optSheet) {
  var sheet = optSheet || SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Orders")
  var products = sheet.getRange(1, PRODUCT_COLUMN, sheet.getLastRow()).getValues()
  return ArrayLib.indexOf(products, 0, product) +1
}

function getPreTweakedProduct(product){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Pre-tweaked Orders")
  return ArrayLib.filterByText(sheet.getDataRange().getValues(), PRODUCT_COLUMN, product)
}

//--------

function getSsNames(){// WORKS but VERY slow - opens all files to get names...
  const files = getSsSortByName("^Dry Orders Merged*")
  log(files.map(sheet => SpreadsheetApp.open(sheet).getName()))
}

function getSsSortByName(searchStr) {// returns sorted array of matching file objects (descending)
  return getSheets(searchStr).sort()
}


function getSheets(searchStr) { // change returns...
  var ss = SpreadsheetApp.getActiveSheet();
  var files = DriveApp.getFilesByType(MimeType.GOOGLE_SHEETS);
  var filesArray = []
 
  while (files && files.hasNext()) {
    var file = files.next();
    if (file.getName().match(searchStr)) {
      filesArray.push(file);
    }
  }
  return filesArray
}


function getLatestSS(optSearchStr){ //actually gets last in alphabetical order
  return getSsSortByName(optSearchStr).slice(-1)[0]
}


function getLatestSsId(optSearchStr){ 
  return getLatestSS(optSearchStr).getFileId()
}
 
//-------------

function say(obj){
  Logger.log(JSON.stringify(obj, null, 4))
}

function log(arg){
  var arr = (arg.constructor === Array) && arg || [arg]
  arr.unshift(new Date())
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("RunLog") ||
               SpreadsheetApp.getActiveSpreadsheet().insertSheet('RunLog').hideSheet()
  sheet.appendRow(arr)
}


function rounded(x) {
  return Math.round(x*100)/100
}


function isDry() {
  var re = /Dry/i;
  return re.test(SpreadsheetApp.getActiveSpreadsheet().getName())  
}

function isFresh() {
  var re = /Fresh/i;
  return re.test(SpreadsheetApp.getActiveSpreadsheet().getName())
}

function isNumeric(n) {
  return !isNaN(parseFloat(n)) && isFinite(n);
}

function dynamicSort(property) {   // eg members.sort(dynamicSort('id'))
    var sortOrder = 1;
    if(property[0] === "-") {
        sortOrder = -1;
        property = property.substr(1);
    }
    return function (a,b) {
        var result = (a[property] < b[property]) ? -1 : (a[property] > b[property]) ? 1 : 0;
        return result * sortOrder;
    }
}

//--------

function reducePenalty(){//correct price from 20% penalty loading to 10% loading
  var cell = SpreadsheetApp.getCurrentCell()
  cell.setValue(cell.getValue() / 1.2 * 1.1)
  cell.offset(1, 0).activateAsCurrentCell()
}

function sharePdfPacksheets(pdf){
  pdf.addViewers([COOP_EMAIL])  
}
