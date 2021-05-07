// ROLLOVER

//        Save id/date/name/balance for each order completed in current orders (run just before rollover)
// v2.01  refreshOrders: Repair code to make units consistent
// v2.00 Removed Fresh code
// v1.99  generalise notify to notify(recipients, subject, msg), special case notifyNico()
// v1.98  addDays: replace call to getYear(was getting 2 digit year)  with getFullYear  (getting 4 digit year)
//                - seems to have changed behaviour a few months ago, with v8? 
// v1.97  Move email addresses to globals
// v1.967 FRESH Replace James and Susannah with Carol Shortis
// v1.966 DRY Before rollover, check members sheet is linked (as well as banking sheet). 
// v1.965 rolloverDates - Change time between releases to 21 days instead of 28
// v1.964 Draft function to release sheet: release()
//        Next: improve html, personalise with account details
// v1.963 AS Phoebe has left, comment out her share being added to her totals
// v1.962 fix test for Tiff/Phoebe - isFresh should be isFRESH
// v1.961 may rollover if date within 7 days, instead of 5
// v1.96 Remove a couple of comments
// v1.95 Change rollover notification recipient - replace Seraphim with Susannah, James. Add Kasey
//       Phoebe and Tiff code got added somewhere in here
// v1.94 if not ok to rollover, Activate sheet that requires fix, Log calls
// v1.93 Correcting daylight saving error, moving validity tests to the front and improving notification

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
  const oldSS = SpreadsheetApp.getActiveSpreadsheet();
  const newSS = oldSS.copy("Fresh auto created test sheet")
  
  SpreadsheetApp.setActiveSpreadsheet(newSS)
  
  clearRolledOver();
  var editors = oldSS.getEditors();
  for (var i = 0; i < editors.length; i++){
    newSS.addEditor(editors[i]);
  }
}


function rollover() {//Rollover order - preparing new sheet
  if (okToRollover()){
    deleteRunLog();
    
    setStatus("Not ready");
    setRolledOver();        // if it falls over and is incomplete, we do not want to be able to run it again without cleaning up the mess

    noteMembersHaveOrdered()
    rolloverTotals(); 
    rolloverDates();

    log('Deletions...')
    deleteOrders();
    deletePreTweakSheet();
    deleteChangeLog();


    refreshFormulae(); 
    addMembers();
    
    rolloverRosters();
    
    notifyNico();
    triggerReminders();
    
    // Remove users
    // Load notices   
    log('Rollover completed successfully')
  }
}  


function rolloverDates(){
  log('rolloverDates...')

  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var rangeNextBD = ss.getRangeByName('tot_Next_Balance_Date')
  var nextBD = new Date(rangeNextBD.getValue())
  
  copyNamedRange('tot_Current_Balance_Date', 'tot_Previous_Balance_Date')            // copy Current to Previous
  copyNamedRange('tot_Next_Balance_Date', 'tot_Current_Balance_Date')                // copy Next to Current
  
  rangeNextBD.setValue(addDays(21, nextBD))                 // add 21 days to Next (was 28)

}

function addDays(numDays, initialDate){
  var minutes = 1000 * 60; // ms/sec * sec/min
  var hours = minutes * 60;
  var days = hours * 24;
  
  var newDateTime = new Date(initialDate.getTime() + numDays*days + 1*hours) // 28 days later between 00:00 - 02:00 depending on daylight saving transitions
  var newDate = new Date(newDateTime.getFullYear(), newDateTime.getMonth(), newDateTime.getDate())  // drop time component
  return newDate
}


function deleteOrders(){
  var ss = SpreadsheetApp.getActiveSpreadsheet().getRangeByName('ord_Data').clearContent();
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
  copyDown(ss.getRangeByName("ord_TotalOrdered"))
  
  // make units consistent... in lowercase and "ea" not "each"
  var range = ss.getRangeByName("ord_Unit")
  var data = range.getValues()
  range.setValues(data.map(x => x[0].match(/each/i) ? ["ea"]
                                                    : [x[0].toString().toLowerCase()]))
}





function matchNumLines(){ // after adding products to Orders, add lines to Totals then refresh totals
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var sheet=ss.getSheetByName("Totals")
  var lastTotal = ss.getRangeByName("tot_Order_Subtotals").getRow() - 1
  var lastOrder = ss.getSheetByName("Orders").getLastRow() - 1
  
  if (lastOrder > lastTotal) {// more orders than totals
    sheet.insertRowsBefore(lastTotal-10, lastOrder-lastTotal)   //add lines to totals
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
    log('rolloverTotals...')
  var ss = SpreadsheetApp.getActiveSpreadsheet();
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

function notifyNico() {
  notify("affordableorganics07@gmail.com", "New Sheet", "New spreadsheet is ready for updating...")
}

function notify(recipients, subject, msg){
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var url = ss.getUrl()
  var ssName = ss.getName()
  var message = {to: IT_EMAIL + ", " + recipients,
                 subject: subject, //+ " - " + ssName,
                 htmlBody: "<a href='" + url + "'>" + ssName + "</a>" + brbr + msg
                }
  MailApp.sendEmail(message)
  //log (["Emailed...", recipients])

}

function tellIT(msg, optUrl){
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var url = optUrl || ss.getUrl()
  var ssName = ss.getName()
  var subject = "Dry Co-op - coded message"
  var message = {to: IT_EMAIL,
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
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var days = 1000 * 60 * 60 * 24  //  num of ms/day
  var ui = SpreadsheetApp.getUi()

  
  // Have we rolled over already?
  var props = PropertiesService.getDocumentProperties()
  var rolledOver = props.getProperty("RolledOver")

  if (rolledOver === "true") {
    ui.alert("Oops, you probably didn't mean to do that - it seems rollover has already been run and running it again might really screw things up.")
    return false
  }

  // Have banking and members workbooks been connected to this workbook? (otherwise value in a1 shows #REF! and the banking transactions are unavailable)
  var bankingSheet = ss.getSheetByName("Banking")
  var bankingOK = (bankingSheet.getRange("A1").getValue() != "#REF!")
  
  var membersSheet = ss.getSheetByName("Members")
  var membersOK = (membersSheet.getRange("J1").getValue() != "#REF!")

  if (!bankingOK) {
    bankingSheet.activate()
    if (membersOK) {
      ui.alert("Oops, can't run the rollover until the banking link has been reconnected.")
    } else {
      ui.alert("Oops, can't run the rollover until banking and members links have been reconnected.")
    return false
    }
  } else {
    if (!membersOK) {
      membersSheet.activate()
      ui.alert("Oops, can't run the rollover until the members link has been reconnected.")
      return false
    }
  }
    
 
 
  // Is the banking rollover date correct? This has to be manually adjusted over summer - IN BOTH the old sheet and the new sheet
  // Closing date for banking for the last ss of the year must run all the way up to the release of the next sheet in Jan/Feb
  
  var closeDate = new Date(ss.getRangeByName('tot_Next_Balance_Date').getValue())
  if (closeDate < Date.now()-7*days){
    ss.getRangeByName('tot_Next_Balance_Date').activateAsCurrentCell()
    ui.alert("Oops, closing banking date should be in the last 7 days. Change date here AND on the previous spreadsheet.")
    return false
  }
  
  return true
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
  
function getPackDateFromFilename(ss = SpreadsheetApp.getActive()){
  var name = ss.getName()
  var re = /\b(19|20\d\d)([-\/])(0[1-9]|1[012])\2(0[1-9]|[12][0-9]|3[01])\b/;      //matches: full date, year,  separator,  month, day
  var matches = re.exec(name)

  return ( new Date(matches[1] ,matches[3]-1, matches[4]) )                       // month is 0 indexed
}

function getNextPackDateFromFilename(){
  var currPackDate = getPackDateFromFilename()
  

}

function getCloseDate(){


}

function triggerReminders() {
  log('triggerReminders...')

  var trigger = ScriptApp.newTrigger("sendReminderSMS")
    .timeBased()
    .inTimezone("Pacific/Auckland")
    .onWeekDay(ScriptApp.WeekDay.SUNDAY)
    .atHour(16)
    .create();    

  trigger.getUniqueId()
}


function rolloverRosters(){
  log('rolloverRosters...')

  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var range = ss.getRangeByName('ros_This_Pack')
  ss.getSheetByName('Roster').hideRows(range.getRow() - 1, range.getNumRows() + 2)
  ss.setNamedRange('ros_This_Pack', range.offset(range.getNumRows()+2, 0))
}

function release(){// draft
  var ss = SpreadsheetApp.getActiveSpreadsheet()

  //openOrdering()
  
  // assemble message
  
  var data = ss.getRangeByName("not_Notices").getDisplayValues()
  var line = ""
  var msg  = "<table>"
  
  for (var r=0; r < data.length; r++) {
    line = "<tr>"
    for (var c=0; c < data[r].length; c++){
      line += "<td>" + data[r][c] + "</td>"
    }
    line += "</tr>"
    msg += line
  }
  msg += "</table>"
  
//  couldn't get this bit to work - try getting image from Photos?
//  msg += brbr + "<img src='https://drive.google.com/open?id=0B8U6153AfrnmTTB5a2dsTlFxUklZTnBUTjY1VzB0Z3gySW1r'>"
  
  var recipients = "affordableorganics07@gmail.com"
//  var recipients = ((isFRESH && "mattrobin24@gmail.com,  matt.mcrae86@gmail.com"
//                    + "") ||
//                    ("affordableorganics07@gmail.com"))
  
  var ssName = ss.getName()
  var link = "<a href='" + ss.getUrl() + "'>" + ssName + "</a>"
  
  var message = {to: IT_EMAIL + ", " + recipients,
                 subject: ssName,
                 htmlBody: msg + brbr + link
                }
  
  // send notification to all members

  MailApp.sendEmail(message)
}

/********
 * 
 * One-off code - step through sheets and log accounts to Dry DB
 *    applies noteMembersHaveOrdered to all post merge sheets
 *    Should only be run to fix problems - delete the existing data first
 */

function getPastAccountBalancesForDB(){  
  const sss = getSsSortByName("^Dry Orders Merged 20")
  sss.map(x => noteMembersHaveOrdered(SpreadsheetApp.open(x)))

}

/***********
 * 
 * Collects data from Totals sheet and adds to PastOrderTotals in Dry DB
 * should be called during rollover, prior to reset of new sheet
 */ 

function noteMembersHaveOrdered(ss = SpreadsheetApp.getActiveSpreadsheet()) {

  const packDate = getPackDateFromFilename(ss)

  const totIDs = ss.getRangeByName("tot_IDs").getDisplayValues()[0]
  const names = ss.getRangeByName("tot_members").getValues()[0]
  const currOrders = ss.getRangeByName("tot_Current_Orders").getValues()[0]
  const currCredits = ss.getRangeByName("tot_Current_Credits").getValues()[0]
  const provBalances = ss.getRangeByName("tot_Provisional_Balances").getValues()[0]  // provBalances are about to become current


  var data = ArrayLib.transpose([totIDs, names, currOrders, currCredits, provBalances])
  
  // dispose of the last row of data as this represents the final formatting column in the totals spreadsheet
  data.pop()

  // select those who have ordered or been credited, and include packDate
  data = data.filter(x => x[2] < -MIN_ORDER_FEE || x[3] !== 0).map(x => [x[0], packDate, x[1], x[2], x[3], x[4]])

  //write data array to PastOrderTotals in DRY DB
  const ssDry = SpreadsheetApp.openById(DB_ID)
  const sheet = ssDry.getSheetByName("PastOrderTotals")
  sheet.getRange(sheet.getLastRow() + 1, 1, data.length, data[0].length)
    .setValues(data)

  // get start and end dates for this pack
  const payStart = ss.getRangeByName("tot_Current_Pay_Start").getValue()
  const payEnd = ss.getRangeByName("tot_Current_Pay_End").getValue()

  const packSheet = ssDry.getSheetByName("Packs")
  const prevID = packSheet.getRange(packSheet.getLastRow(), 1).getValue()
  const newID = "P"+ (parseInt(prevID.match(/\d+/)) +1)
  packSheet.appendRow([newID, packDate, payStart, payEnd])
}
