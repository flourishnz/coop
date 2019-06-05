// MEMBERS
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
  
  this.fullName = function() {
    return this.firstName + " " + this.lastName
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



function getMembers(){// returns array of objects - needs fix returning blank lines at the bottom...
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var data = ss.getRangeByName("mem_Data").getValues()
  var members = []
  var id
  
  for (var i = 1; i < data.length; i++){
    if (isDRY){
      id = data[i][0].toString()
      members.push({id: id ,
                    firstName: data[i][1],
                    lastName: data[i][2],
                    name: data[i][1] + " " + data[i][2],
                    mobile: data[i][3].toString(),
                    email: data[i][4],
                    otherPhone: data[i][5].toString(),
                    gmailAccounts: data[i][6],
                    homeAddress: data[i][7],
                    row: i
                   })
    } else {
      id = data[i][1].toString()
      members.push ({role: data[i][0],
                     id: id,
                     name: data[i][2],
                     email: data[i][3],
                     homePhone: data[i][4].toString(),
                     mobile: data[i][5].toString(),
                     homeAddress: data[i][6] + (data[i][7] ? (', ' + data[i][7]) : ''),
                     row: i
                   })
    }
  }
  return members   // an array of objects
}



function getMember(arg){// arg is Id or row number in Members Tab(as reported by onEdit)
  var sheet = SpreadsheetApp.getActive().getSheetByName("Members") 
  var data = sheet.getDataRange().getValues()
  var member = new Member();
  var id = isValidId(arg) && arg || isNumeric(arg) && arg <= data.length && data[arg-1][MEM_ID_OFFSET]
  var i = ArrayLib.indexOf(data, MEM_ID_OFFSET, id)   // look for id
  if (i<0)  {
    log(['Member not found in Members tab', id])
    SpreadsheetApp.getUi().alert('Member not found in Members tab: ' + id);
    return {}}
  
  if (isDRY){
    member.id = id
    member.firstName = data[i][1]
    member.lastName = data[i][2]
    member.name = data[i][1] + " " + data[i][2]
    member.mobile = data[i][3].toString()
    member.email = data[i][4]
    member.otherPhone = data[i][5].toString()
    member.gmailAccounts = data[i][6]
    member.homeAddress = data[i][7]                                    
    member.row = i
  } else {
    // Fresh
    member.role = data[i][0]
    member.id = id
    member.name = data[i][2]
    member.email = data[i][3]
    member.otherPhone = data[i][5].toString()
    member.mobile = data[i][5].toString()
    member.homeAddress = data[i][6] + (data[i][7] ? (', ' + data[i][7]) : '')                                 
    member.row = i
  }

  return member
}

//------------------------------------------------------------------------


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
  if (!_.isEmpty(member)) {
    saveExMemberDetails_(member)
    removeFromOrders_(member)
    removeFromTotals_(member)
    removeFromMembers_(member)
    //  removeFromCurrentContacts(member) //must be actioned by coop account
    // revoke access?
    // send an email to Seraphim/Joanne/coop/kasey etc
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
