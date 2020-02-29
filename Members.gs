// MEMBERS

// STILL TO DO...: removeMember - remove from contacts, revoke access

// v2.3  Rewrite getMember and getMembers to use Member class and to share code
//       Also adding in more name handling
// v2.2  Notify everyone who needs to know when somone is removed from the co-op
//       Remove from Contacts
// v2.1 add getPreMergeMember - so that (Dry) members in Orders/Totals sheets but
//        no longer in Members sheet can be properly deleted
//       Commented out alert to say Member wasn't found but still logging to runlog...
// put this back???
//       Commented out row because a) it appeared to be wrong/misleading
//         b) wasn't using it
// v2.01 removeFromCurrentContacts has been moved to Contacts.gs
// v2.0 Adjust for new Dry Members sheet layout after Merge
// v1.5.1 Logging call to addMembers

// 11-8-18 added code to remove member from spreadsheet - still needs some refinement
// 15-3-18 fixed bug ss not defined in function Member and SYNCHED
// 1-3-18 synched - remove member OR NOT
// 8-12-17 synched Fresh and Dry b
// 28-11-17 add multiple new members from Members sheet
// 16-11-17 synched with dry a 
// 30-9-17 Fresh code - completing steps
// 29-9-17 B code - synched with Fresh - commented code at the bottom still different
//changes...
//   


function Member() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  this.getFullName = function() {
    return (isDRY ? this.firstName + " " + this.lastName
                 : this.name)
  }  
  
  this.getCurrentBalanceDate = function() {
    return ss.getRangeByName("tot_Current_Balance_Date").getValue()
  }

  this.getCurrentBalance = function(){
    var balances = ss.getRangeByName("tot_Current_Balances").getValues()
    var ids = ArrayLib.transpose(ss.getRangeByName("tot_IDs").getValues())
    var i = ArrayLib.indexOf(ids, 0, this.id)
    if (i) {
      this.col = i
      return balances[0][i]
    }
  }
 
  this.getPayments = function() {
    return getLatestTransactions(this.id)
  }
 
  this.getLatestPayment = function() {
    var transactions = getLatestTransactions()
    if (this.id in transactions) {
      return transactions[this.id][0]
    } else {
      return {date: '', payment: 0}
    }
  }
  
}

function addMembers() {
  log('addMembers...')

  var newIDs = getNewMemberIDs()
 
  for (var i = 0; i < newIDs.length; i++) {
    addMember(getMember(newIDs[i]))        // getting new member details from Members sheet
  }
}



function addMember(member) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); 
  var totals = ss.getSheetByName("Totals");
     
  // add to contacts
  //addMemberToContacts(member) TRYING to get this to happen automagically on editing of Members sheet
  
  // add to Totals sheet
  var tcol = insertColumn(totals)
  totals.getRange(ss.getRangeByName("tot_Members").getRow(), tcol).setValue(member.name)
  totals.getRange(ss.getRangeByName("tot_IDs").getRow(), tcol).setValue(member.id)
  log(["Added member to totals", member.id, member.name])
  
  // ... and initialise balances  
  totals.getRange(ss.getRangeByName("tot_Previous_Balances").getRow(), tcol).setValue(0)
  totals.getRange(ss.getRangeByName("tot_Previous_Orders").getRow(), tcol).setValue(0)
  totals.getRange(ss.getRangeByName("tot_Previous_Credits").getRow(), tcol).setValue(0)
  
  // ... and charge membership fee
  totals.getRange(ss.getRangeByName("tot_Current_Credits").getRow(), tcol).setValue(-MEMBERSHIP_FEE).setNote("Membership fee")
  
  //   add to Orders sheet
  insertColumn(ss.getSheetByName("Orders"))
  log(["Added member to orders", member.id, member.name])
  
  //   share worksheet
  shareSheet(member)
}


function shareSheet(member){
  var ss = SpreadsheetApp.getActiveSpreadsheet(); 
  var accounts = findEmailAddresses(member)

  for (i in accounts){
    try {
      ss.addEditor(accounts[i])
      log(["Shared sheet with", member.name, accounts[i]])
    }
    catch (err) {
      SpreadsheetApp.getUi().alert("Invalid email address: " + accounts[i])
      log(["Invalid email address ", member.name, accounts[i]])
    }
  }
}

function findEmailAddresses(member) {
  var pattern = /[a-zA-Z0-9._-]+@[a-zA-Z0-9._-]+\.[a-zA-Z0-9._-]+/gi
  return  member.gmailAccounts.match(pattern) || member.email.match(pattern) || []
}


function getNewMemberIDs(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  // get newest in Members sheet
      // var newestID =  ss.getSheetByName("Members").getDataRange().getValues().pop()[MEM_ID_OFFSET]
  var memIDs = ss.getRangeByName("mem_IDs").getValues()
  memIDs.shift()  // remove header
  var newestID = Math.max.apply(Math, memIDs)

  var existingIDs = ss.getRangeByName("tot_IDs").getValues()     // already  in Totals sheet
  var lastID = Number(existingIDs[0][existingIDs[0].length-2])   // -1 for empty colum and -1 to get offset instead of count
  var newIDs = []
  
  for (var id = lastID+1; id <= newestID; id++){
    newIDs.push(id.toString())
  }     
  return newIDs
}


function insertColumn(sheet) {
  var numRows = sheet.getLastRow()
  var col = sheet.getLastColumn()
  
  //.. insert 1 column before last column
  sheet.insertColumnBefore(col)            // after insertion col is the place of the new column

  //.. copy width, formulae and format from previous column
  sheet.setColumnWidth(col, sheet.getColumnWidth(col-1))
  sheet.getRange(1, col-1, numRows).copyTo(sheet.getRange(1, col, numRows))
  sheet.getRange(1, col, numRows).clearNote()
  
  // if Orders, clear values
  if (sheet.getName() == "Orders") {
    sheet.getRange(FIRST_ORDER_ROW , col, numRows-FIRST_ORDER_ROW+1).clearContent()
  }
  
  return col
}



function getMembers(){// returns array of objects
  var data = SpreadsheetApp.getActiveSpreadsheet()
               .getRangeByName("mem_Data")
               .getValues()
               .filter(function(row){return isValidId(row[MEM_ID_OFFSET])})  // drop headers and any invalid member (ids)
  return data.map(toMember)  
}

function getMember(arg){// arg is Id or row number in Members Tab(as reported by onEdit) no longer checking for this?
  var sheet = SpreadsheetApp.getActive().getSheetByName("Members") 
  var data = sheet.getDataRange().getValues()
  
  // locate member, if there
  var id = isValidId(arg) && arg || isNumeric(arg) && arg <= data.length && data[arg-1][MEM_ID_OFFSET]
  var i = ArrayLib.indexOf(data, MEM_ID_OFFSET, id)   // look for id
  if (i<0)  {
    log(['Member not found in Members tab', id])
    // SpreadsheetApp.getUi().alert('Member not found in Members tab: ' + id);
    return {}
  }
  
  return toMember(data[i])
}

function toMember(row) {
  var member = new Member()
  var matches
  if (!isValidId(row[MEM_ID_OFFSET])) {return {}}
  
  if (isDRY){    
    member.id = row[0].toString()
    member.firstName = row[1]
    member.lastName = row[2]
    member.name = row[1] + " " + row[2]
    member.mobile = row[3].toString()
    member.email = row[4]
    member.otherPhone = row[5].toString()
    member.gmailAccounts = row[6]
    member.homeAddress = row[7]                                    
    members.push (member)           
  } else {
    // Fresh
    member.role = row[0]
    member.id = row[1].toString()
    
    member.name = row[2].toString().trim()      
    matches = /(\w+[-']?\w*)\W+(.*)/.exec(member.name)
    member.firstName = matches[1]
    member.lastName = matches[2]

    member.email = row[3]
    member.otherPhone = row[4].toString()
    member.mobile = row[5].toString()
    member.homeAddress = row[6] + (row[7] ? (', ' + row[7]) : '')                                 
  }

  return member
}


function getPreMergeMember(id){// only Dry has this sheet - returns {} to Fresh
  log(['Entered getPreMergeMember'])
  try {
    var sheet = SpreadsheetApp.getActive().getSheetByName("preMergeMembers")
    var data = sheet.getDataRange().getDisplayValues()
    }
  catch (e) {
    log(['preMergeMembers Sheet not found.'])
    return {}
  }
  
  var member = new Member();
  var i = ArrayLib.indexOf(data, 0, id)   // look for id
  if (i<0)  {
    log(['Member not found in preMergeMembers tab', id])
    //SpreadsheetApp.getUi().alert('Member not found in preMergeMembers tab: ' + id);
    return {}
  }
  
  member.id = id
  member.name =  data[i][2]
  member.mobile = data[i][5].toString()
  member.email = data[i][3]
  member.otherPhone = data[i][4].toString()

  return member
}

//------------------------------------------------------------------------

function testMe(){
  say (getMembers())
}

function removeThisMember(){
  // call from Totals sheet or from Members sheet to initiate removal of 'active' member
  var range = SpreadsheetApp.getActiveRange();
  var sheet = range.getSheet()
  var sheetName = sheet.getName()
  var thisCol = range.getColumn()
  var thisRow = range.getRow()
  var ui = SpreadsheetApp.getUi();

  if (sheetName == 'Totals'){
    if (thisCol <= 4){
      ui.alert("Please move to a member column.")
    } else {
      var response = ui.alert("Remove " +  sheet.getRange(TOT_ID_ROW-1, thisCol).getValue() + " from the co-op?", ui.ButtonSet.YES_NO)
      if (response == ui.Button.YES) {removeMember(sheet.getRange(TOT_ID_ROW, thisCol).getValue())}
    }
    return
  } 
  else if (sheetName == 'Members'){
    if (thisRow == 1){
      ui.alert("Please move to a member row.")
      return
    }
    var response = ui.alert("Remove " +  sheet.getRange(thisRow, MEM_ID_OFFSET+1).getValue() + " " +
        sheet.getRange(thisRow, MEM_ID_OFFSET+2).getValue() + " " +
        sheet.getRange(thisRow, MEM_ID_OFFSET+3).getValue() +
        " from the co-op?", ui.ButtonSet.YES_NO)
        if (response == ui.Button.YES) {removeMember(sheet.getRange(thisRow, MEM_ID_OFFSET+1).getValue())}
  } 
  else {ui.alert("Select a cell in the member's row in Members or in the member's column in Totals and try again.")}
}



function removeMember(id) {
  var member = getMember(id)
  if (_.isEmpty(member)) {
    member = getPreMergeMember(id)
  }
  
  if (_.isEmpty(member)) {
    SpreadsheetApp.getUi().alert(id + ' not found in Members or preMergeMembers sheets\n' +
                                  'Member not removed.')
  } else {
    notifyRemoval(member)
    
    saveExMemberDetails_(member)
    removeFromOrders_(member)
    removeFromTotals_(member)
    removeFromMembers_(member)
//    removeFromCurrentContacts(member) //must be actioned by coop account
    // revoke access?
  }
}



function saveExMemberDetails_(member){// still needs refining...
  var sheet = SpreadsheetApp.getActive().getSheetByName('Ex Members')
  var newRow = sheet.getLastRow() -1
  sheet.insertRowAfter(newRow-1)    // inserted row is now newRow

  if (isDRY){
    var data = [[member.name, member.id, member.getCurrentBalanceDate(), member.getCurrentBalance(),
                      '', '', '', '', '', '', '',
                      member.id, member.name, member.email, member.mobile , member.homePhone]]
    
    sheet.getRange(newRow, 2, 1, data[0].length)
         .setValues(data)
    
    // add formulae
    sheet.getRange(newRow,  7).setFormula('=sumifs(bank_Amount, bank_Bin, ex_ID, bank_Date, ">" & ex_Date, bank_Amount, ">0")')     //deposits
    sheet.getRange(newRow, 8).setFormula('=sumifs(bank_Amount, bank_Bin, ex_ID, bank_Date, ">" & ex_Date, bank_Amount, "<0")')      //refunds
    sheet.getRange(newRow, 10).setFormulaR1C1('sum(R[0]C[-5]:R[0]C[-1])')                                                           //net balance

  } else {//FRESH
    var data = [[member.name, member.id,  '', '', member.getCurrentBalanceDate(), member.getCurrentBalance(),
                     50, '', '', '', '', '', '', 
                     member.id, member.name, member.email, member.mobile , member.homePhone, member.homeAddress]]
    
    sheet.getRange(newRow, 2, 1, data[0].length)
         .setValues(data)
    
    // add formulae
    sheet.getRange(newRow,  9).setFormulaR1C1('=sumifs(bank_Amount, bank_Bin, ex_ID, bank_Date, ">" & ex_Date, bank_Amount, ">0")')  //deposits
    sheet.getRange(newRow, 10).setFormulaR1C1('=sumifs(bank_Amount, bank_Bin, ex_ID, bank_Date, ">" & ex_Date, bank_Amount, "<0")')  //refunds
    sheet.getRange(newRow, 12).setFormulaR1C1('sum(R[0]C[-5]:R[0]C[-1])')                                                            //net balance
  }


}


function notifyRemoval(member, optUrl){
  var subject = member.getFullName() + " has left the " + (isFRESH ? "Fresh" : "Dry") + " co-op"
  member.leavingBalance = Number(member.getCurrentBalance()) + MEMBERSHIP_BOND
  var details = formatAcctDetails_(member)
  
  var recentPayments = formatPayments_(member)
  var link = formatLink(optUrl)
  var autoGenMsg = brbr + "<small>This message was automatically generated. Please contact " +
                    IT_NAME + " at " +  IT_EMAIL + " if you have any queries."; 
  
  var ui = SpreadsheetApp.getUi()
  

  if (isDRY) {
    MailApp.sendEmail({
      to: [IT_EMAIL, COOP_EMAIL].join(',') ,
      subject: subject,
      htmlBody: details + recentPayments + autoGenMsg
    })
  } else {// isFRESH
    // Co-op closed
    var closure = "At the recent AGM of Kapiti Fresh co-op, the members reluctantly voted to wind up the co-op." + brbr
    
    // notify Member, Treasurer and IT
    var action = (member.leavingBalance > 0 
      ? " Please forward your account details to " + TREASURER_EMAIL + " so that " + TREASURER_NAME + " can arrange a refund." + brbr
      : "Please contact our treasurer "
         + TREASURER_NAME + " at " + TREASURER_EMAIL + " if you wish to make special payment arrangements." + brbr);
  
    MailApp.sendEmail({
      to:[member.email, IT_EMAIL, TREASURER_EMAIL].join(','),
      subject: "Fresh co-op account deleted - " + member.getFullName(),
      htmlBody: "Hi " + member.firstName + brbr 
      + closure
      + "Your " 
      + (isFRESH ? "Fresh" : "Dry") 
      + " co-op account has been deleted. Your net balance is $" 
      + Math.abs(member.leavingBalance).toFixed(2) 
      + (member.leavingBalance<0 ? " in debit." : " in credit.")
      + brbr
      + action
      + details
      + recentPayments
      + autoGenMsg
    })
   
 
//    // notify Rosters officer
//    MailApp.sendEmail({
//      to:ROSTERS_EMAIL,
//      subject: subject,
//      htmlBody: "Hi " + ROSTERS_NAME + brbr
//         + "Please remove " + member.firstName + " (" + member.id + ") from the rosters."
//         + autoGenMsg
//    })
  
//    // notify Membership officer
//    MailApp.sendEmail({
//      to:MEMBERSHIP_EMAIL,
//      subject: "FYI: " + subject,
//      htmlBody: autoGenMsg
//    })
    
  } //isFresh
}

function formatAcctDetails_(member){
  var date = Utilities.formatDate(member.getCurrentBalanceDate(), "GMT+12:00", "d MMMM yyyy")
  var details = "<table>"
  details += "<tr><td>ID</td><td>" + member.id + "</td></tr>"
  details += "<tr><td>Name</td><td>" + member.getFullName() + "</td></tr>"
  details += "<tr><td>Balance Date</td><td>" + date + "</td></tr>"
  
  details += (isFRESH ? "<tr><td>Closing Balance</td><td>" + "$" + Math.abs(member.getCurrentBalance()).toFixed(2)
                                                         + (member.getCurrentBalance()>=0 ? " in credit" 
                                                                                          : " in debit")

                        + "</td></tr>"
                        + "<tr><td>Bond Refund</td><td>$50.00</td></tr>" 
                      : "")
  details += "<tr><td>Net Balance</td><td>" + "$" + Math.abs(member.leavingBalance).toFixed(2)
                                                  +(member.leavingBalance>=0 ? " in credit" 
                                                                           : " in debit")
                                                  + "</td></tr>"
  details += "</table>"       

  return details
}

function formatLink(optUrl){
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var url = arguments.length = 1 && optUrl || ss.getUrl()
  return brbr + "<a href='" + url + "'>" + ss.getName() + "</a>"
}

function formatPayments_(member) {
  var payments = member.getPayments()[member.id].reverse()
  var html = brbr + "<h3>Your Most Recent Payments</h3><table>"
  html += payments.reduceRight(function (h, payment){
        return h + "<tr><td>" + formatDate(payment.date) + "</td><td>" 
                 + "$" + payment.amount.toFixed(2) + "</td></tr>"
  }, "")
  html += "</table>"
  return html
}

function formatDate(date){
  return Utilities.formatDate(new Date(date), "GMT+12:00", "d MMM yyyy")
}

function removeFromMembers_(member){
  var sheet = SpreadsheetApp.getActive().getSheetByName("Members")
  var data = sheet.getDataRange().getValues() 
  var i = ArrayLib.indexOf(data, MEM_ID_OFFSET, member.id)   // look for id
  if (i == -1)  {
    log(["Couldn't remove member from Members sheet", member.id])
    return 
  }
  sheet.deleteRow(i+1)
  log(['Removed member from Members tab', member.id, member.name])
}


function removeFromOrders_(member){  // remove from Totals and Orders sheets
  var ss = SpreadsheetApp.getActive()
  var ids = ss.getRangeByName('ord_Bins').getValues()[0]
  var col = ids.indexOf(member.id)
  if (col == -1) {
    log(['Failed to remove member - member not found on Orders sheet', member.id, member.name])
    return
  }
  say(ids)
  if (col > -1) {
    ss.getSheetByName('Orders').deleteColumn(col+1)
    log(['Removed member from Orders', member.id, member.name])
  } else {
    log(['Failed to locate ' + member.id + ' on Orders sheet.'])
  }
}

  
function removeFromTotals_(member){
  var ss = SpreadsheetApp.getActive()
  var ids = ss.getRangeByName('tot_Bins').getValues()[0]
  var col = ids.indexOf(member.id)
  say(ids)
  if (col == -1) {
    log(['Failed to properly remove member - member not found on Totals sheet', member.id, member.name])
    return
  }
  ss.getSheetByName('Totals').deleteColumn(col+1)
  log(['Removed member from Totals', member.id, member.name])
}

//function removeFromCurrentContacts(member) - this code has gone to Contacts

function getLatestTransactions() {// saves to ScriptProperties
  // if transactions have already been cached...
  var cached = CacheService.getScriptCache().get('transactions')
  if (cached !== null) {return JSON.parse(cached)}

  // otherwise generate from sheet
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var data = ss.getSheetByName("Banking").getDataRange().getValues()
  var idCol = ss.getRangeByName("bank_Bin").getColumn()-1
  var amountCol = ss.getRangeByName("bank_Amount").getColumn()-1
  var transactions = []

  for (var i = 0; i < data.length; i++) {
    var id = data[i][idCol]
    if (!isValidId(id)) {continue}
    
    if (!(id in transactions)) {transactions[id] = []}
    transactions[id].unshift({date: data[i][0], 
                              amount: data[i][amountCol]
                             })
    if (transactions[id].length == 4) transactions[id].pop()
  }

  // then add to cache and return
  CacheService.getScriptCache().put("transactions",  JSON.stringify(transactions))
  return transactions
}