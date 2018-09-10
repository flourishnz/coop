// ROLLOVER
// v1.92

// 14-08-18 Build trigger for reminders
// July 18  rollover rosters
// 4-Jun-18 added seraphim to notifications
// 29-Nov-17 SYNCHED
// 19-Sep-17 SYNCHED fresh - dry AB - 
//         ?  add ord_prices, ord_totalkgs, ord_totalcrates to SS
//    6-Oct 16  check ok to rollover
//   30-July 16 rollover all dates
//   25-July 16 fix syntax error
//   16 July 16 clear notes from rollover totals

function createOrderSheet(){// developing... this code may not run from within a sheet - needs to be in lib
  var oldSS = SpreadsheetApp.getActiveSpreadsheet();
  var newSS = oldSS.copy("Fresh auto created test sheet")
  
  SpreadsheetApp.setActiveSpreadsheet(newSS)
  
  clearRolledOver();
  var editors = oldSS.getEditors();
  for (var i = 0; i < editors.length; i++){
    newSS.addEditor(editors[i]);
  }
}


function rollover() {//Rollover order - preparing new sheet
  if (okToRollover()){
    //var ss = SpreadsheetApp.getActiveSpreadsheet();
    // or, better still, copy sheet
    deleteRunLog();
    
    setStatus("Not ready");
    setRolledOver();        // if it falls over and is incomplete, we do not want to be able to run it again without cleaning up the mess
    rolloverTotals(); 
    rolloverDates();

    deleteOrders();
    if (isFRESH){ clearFreshDirect()  };
    deletePreTweakSheet();
    deleteChangeLog();


    refreshFormulae(); 
    addMembers();
    
    rolloverRosters();
    
    notify("Spreadsheet is ready for updating.")
    triggerReminders();
    
    // Remove users
    // Load notices   
    // Load local
    // Load FreshDirect
    
  }
}  


function rolloverDates(){
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var rangeNextBD = ss.getRangeByName('tot_Next_Balance_Date')
  var nextBD = new Date(rangeNextBD.getValue())

  // this test MAY NOT WORK but not urgent to fix - needed when pack timing changes   - also improve handling
  if (nextBD < Date()-7){
    // Alert Uh-oh This needs to be adjusted on the original ss before creating new sheet - there is a greater than normal interval between packs
    return false
  }
  
  copyNamedRange('tot_Current_Balance_Date', 'tot_Previous_Balance_Date')            // copy Current to Previous
  copyNamedRange('tot_Next_Balance_Date', 'tot_Current_Balance_Date')                // copy Next to Current
  
  if (isDRY) {                                                                     
    rangeNextBD.setValue(new Date(nextBD.getTime()+28*(24*3600*1000)))               // add 28 days to Next
  } else {
    rangeNextBD.setValue(new Date(nextBD.getTime()+14*(24*3600*1000)))               // or add 14 days to Next
  } 
}


function deleteOrders(){
  var ss = SpreadsheetApp.getActiveSpreadsheet().getRangeByName('ord_Data').clearContent();
}


function clearFreshDirect(){
  SpreadsheetApp.getActiveSpreadsheet().getRangeByName("ord_FreshDirect_Pricelist").clearContent()
}
  
function refreshVendors(){// only refreshing main orders at the moment - overwrites chantal over all vendors
  if (isFRESH){
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var prices = ss.getRangeByName("ord_FreshDirect_Pricelist")
    var sheet = ss.getSheetByName("Orders");
    var range = sheet.getRange(prices.getRow(), VENDOR_COLUMN-1, prices.getNumRows(), 2);  //get VENDOR_COLUMN and the one before it (section labels)
    var data = range.getValues()
    
    for (var i=0; i < data.length; i++){
      if (data[i][0] === "") {
        data[i][1] = "Chantal"
      }
    }
    range.setValues(data)
  }
}



function deletePreTweakSheet(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets()
  var re = /^pre-*tweak(ed)? orders$/;
  
  for (var i = 0; i < sheets.length; i++) {
    if (re.test(sheets[i].getName().toLowerCase())) {
      ss.deleteSheet(sheets[i]);
      return;
    }
  }
}


function refreshFormulae() {
  //refreshVendors()    // orders sheet -not running because it just puts chantal over all vendors
  refreshOrders()     // orders sheet
  refreshProducts()   // totals sheet
  refreshTotals()     // totals sheet
}

// Reset formulae in Totals - set all to same formula
function refreshTotals() {
  matchNumLines();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var rangeF = ss.getRangeByName("tot_Formulae");
  rangeF.setFormula("=tot_Prices * hlookup(tot_IDs, ord_Orders, row()-3, false)")
}

function refreshOrders() {
  // reset formulae
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  copyDown(ss.getRangeByName("ord_Prices"))
  if (isFRESH) {
    copyDown(ss.getRangeByName("ord_TotalKgs"))  
    copyDown(ss.getRangeByName("ord_TotalCrates"))
  } else {
    copyDown(ss.getRangeByName("ord_TotalOrdered"))  
  }
  
  range = ss.getRangeByName("ord_Unit")
  units = range.getValues()
  for (i=0; i<units[0].length; i++) {
    if ((units[0][i] ==="each") ||
        (units[0][i] ==="EACH"))
    units[0][i] = "ea"
  }
  range.setValues(units)
}




function matchNumLines(){ // after adding products to Orders, add lines to Totals then refresh totals
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var sheet=ss.getSheetByName("Totals")
  var lastTotal = ss.getRangeByName("tot_Order_Subtotals").getRow() - 1
  var lastOrder = ss.getSheetByName("Orders").getLastRow() - 1
  
  if (lastOrder > lastTotal) {// more orders than totals
    sheet.insertRowsBefore(lastTotal-1, lastOrder-lastTotal)   //add lines to totals
  } else if (lastOrder < lastTotal){
    sheet.deleteRows(10, lastTotal-lastOrder)                  // remove lines from totals
  }
}





function copyDown(range){ // copy formula(e) from first row down the whole range
  var formulae = range.getFormulasR1C1()
  for (var i= 1; i < formulae.length; i ++) {
    formulae[i] = formulae[0]
  };
  range.setFormulasR1C1(formulae)
}


function refreshProducts() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  copyDown(ss.getRangeByName("tot_Products"))
  copyDown(ss.getRangeByName("tot_Prices"))
}


function rolloverTotals() { // Copy curr order details to prev order, starting with orders
  copyNamedRange("tot_Current_Balances", "tot_Previous_Balances");
  copyNamedRange("tot_Current_Orders", "tot_Previous_Orders");
  copyNamedRange("tot_Current_Credits", "tot_Previous_Credits");
  SpreadsheetApp.getActiveSpreadsheet().getRangeByName("tot_Current_Credits").clearNote().clearContent();
}


//Copy source values and notes to destination using named ranges
function copyNamedRange(source, dest) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourceRange = ss.getRangeByName(source);  
  var destRange = ss.getRangeByName(dest);
  
  destRange.setValues(sourceRange.getValues());
  destRange.setNotes(sourceRange.getNotes());
}

function deleteChangeLog(){// clear from second row to end, keep formatting
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Change Log")
  sheet.getRange(2, 1, sheet.getLastRow(), sheet.getLastColumn()).clearContent().clearNote()
}

function deleteRunLog(){
  try {
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("RunLog").clear()
  } 
  catch (error) {// if sheet ain't there we don't care }
  }
}

function notifyNow() {
  notify("New spreadsheet is ready for release...")
}

function notify(msg){
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var recipients = ((isFRESH && "michele@rimuchiro.co.nz ,  matt.mcrae86@gmail.com, iamseraphim@gmail.com") ||
                    ("affordableorganics07@gmail.com"))
  var url = ss.getUrl()
  var ssName = ss.getName()
  var message = {to: "flourish.nz@gmail.com" + ", " + recipients,
                 subject: "New sheet - " + ssName,
                 htmlBody: msg + "<br><br><a href='" + url + "'>" + ssName + "</a>"
                }
  MailApp.sendEmail(message)
}

function tellJulie(msg, optUrl){
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var url = optUrl || ss.getUrl()
  var ssName = ss.getName()
  var subject = ((isFRESH && "Fresh - coded message") || ("Dry - coded message"))
  var message = {to: "flourish.nz@gmail.com",
                 subject: subject,
                 htmlBody: msg + "<br><br><a href='" + url + "'>" + ssName + "</a>"
                }
  MailApp.sendEmail(message)  
}
//---------------------------------------------------------------------------
function setRolledOver(){
  PropertiesService.getDocumentProperties().setProperty("RolledOver", "true")
}

function clearRolledOver(){
  PropertiesService.getDocumentProperties().setProperty("RolledOver", "false")
}

function okToRollover(){
  // return false if already rolled over
  var props = PropertiesService.getDocumentProperties()
  var rolledOver = props.getProperty("RolledOver")
  if (rolledOver === "true") return false

  // if not rolled, return false if banking workbook has not been connected to this workbook (value in a1 is  #REF!)
  // (and therefore the totals are not showing correctly)
  var bankingOK = (SpreadsheetApp.getActiveSpreadsheet()
                   .getSheetByName("Banking").getRange("A1")
                   .getValue() != "#REF!")
  
  return (bankingOK)
}

function reportRolloverStatus(){
  if (okToRollover()){
    Logger.log("ok")
  }
  else {
    Logger.log("Not ok")
  }
}

//--------------------------------------
  
function getPackDateFromFilename(){
  var name = SpreadsheetApp.getActiveSpreadsheet().getName()
  var re = /\b(19|20\d\d)([-\/])(0[1-9]|1[012])\2(0[1-9]|[12][0-9]|3[01])\b/;      //matches: full date, year,  separator,  month, day
  var matches = re.exec(name)

  return ( new Date(matches[1] ,matches[3]-1, matches[4]) )                       // month is 0 indexed
}

function getNextPackDateFromFilename(){
  var currPackDate = getPackDateFromFilename()
  

}


function triggerReminders() {
  if (isFRESH) {
    var trigger = ScriptApp.newTrigger("sendReminderSMS")
    .timeBased()
    .inTimezone("Pacific/Auckland")
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(18)
    .create();
  } else {
    var trigger = ScriptApp.newTrigger("sendReminderSMS")
    .timeBased()
    .inTimezone("Pacific/Auckland")
    .onWeekDay(ScriptApp.WeekDay.SUNDAY)
    .atHour(16)
    .create();    
  }
  trigger.getUniqueId()

}

function rolloverRosters(){
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  if (isFRESH){
    ss.setNamedRange('Roster_This_Pack', ss.getRangeByName('Roster_This_Pack').offset(0,1))
    ss.setNamedRange('Roster_Next_Pack', ss.getRangeByName('Roster_Next_Pack').offset(0,1))
  } else {
    var range = ss.getRangeByName('ros_This_Pack')
    ss.getSheetByName('Roster')
      .hideRows(range.getRow()-1, range.getNumRows()+2)
    ss.setNamedRange('ros_This_Pack', ss.getRangeByName('ros_This_Pack').offset(24,0))
  }
}
