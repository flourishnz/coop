// MEMBERS
// v1.5

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

//  var IDs = ss.getRangeByName("tot_Bins").getValues()[0]
//  var names = ss.getRangeByName("tot_Members").getValues()[0]
//  var prevOrders = ss.getRangeByName("tot_Previous_Orders").getValues()[0]
//  var currOrders = ss.getRangeByName("tot_Current_Orders").getValues()[0]


function Member() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();   
  
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
  
}

function addMembers() { 
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
  
  
  // ... and initialise balances  
  totals.getRange(ss.getRangeByName("tot_Previous_Balances").getRow(), tcol).setValue(0)
  totals.getRange(ss.getRangeByName("tot_Previous_Orders").getRow(), tcol).setValue(0)
  totals.getRange(ss.getRangeByName("tot_Previous_Credits").getRow(), tcol).setValue(0)
  
  // ... and charge membership fee
  totals.getRange(ss.getRangeByName("tot_Current_Credits").getRow(), tcol).setValue(-MEMBERSHIP_FEE).setNote("Membership fee")
  
  //   add to Orders sheet
  insertColumn(ss.getSheetByName("Orders"))
  
  //   share worksheet
  try {
    ss.addEditor(member.email.trim())  
  }
  catch (err) {
    SpreadsheetApp.getUi().alert("Invalid email address: " + member.email.trim())
    Logger.log("Invalid email address: " + member.email.trim() + "\n Worksheet not shared.")
  }  
}




function getNewMemberIDs(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  // get newest in Members sheet
  var newestID =  ss.getSheetByName("Members").getDataRange().getValues().pop()[MEM_ID_OFFSET]
      //ss.getRangeByName("mem_ID").getValues() //.pop()[0]
  

  var existingIDs = ss.getRangeByName("tot_IDs").getValues()     // already  in Totals sheet
  var lastID = Number(existingIDs[0][existingIDs[0].length-2])
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
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var data = ss.getRangeByName("mem_Data").getValues()
  var members = []
  var id
  
  for (var i = 1; i < data.length; i++){
    if (isDRY){
      id = data[i][0].toString()
      members.push({id: id ,
                    name: data[i][2],
                    email: data[i][3],
                    home_phone: data[i][4].toString(),
                    mobile_phone: data[i][5].toString(),
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
  if (i<0)  {return {}}
  
  if (isDRY){
    member.id = id
    member.name = data[i][2]
    member.email = data[i][3]
    member.homePhone = data[i][4].toString()
    member.mobile = data[i][5].toString()
    member.row = i
  } else {
    // Fresh
    member.role = data[i][0]
    member.id = id
    member.name = data[i][2]
    member.email = data[i][3]
    member.homePhone = data[i][4].toString()
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
  
  //else
  if (sheetName == 'Members'){
    if (thisRow == 1){
      ui.alert("Please move to a member row.")
      return
    }
    var response = ui.alert("Remove " +  sheet.getRange(thisRow, MEM_ID_OFFSET+1).getValue() + " " +
                             sheet.getRange(thisRow, MEM_ID_OFFSET+2).getValue() +
                            " from the co-op?", ui.ButtonSet.YES_NO)
    if (response == ui.Button.YES) {removeMember(sheet.getRange(thisRow, MEM_ID_OFFSET+1).getValue())}
  }
}



function removeMember(id) {
  var member = getMember(id)
  saveExMemberDetails_(member)
  removeFromOrders_(member)
  removeFromTotals_(member)
  removeFromMembers_(member)
  //  removeFromCurrentContacts(member) //must be actioned by coop account
  // unshare?
  // revoke access?
  // send an email to Seraphim/Joanne/coop/kasey etc
}



function saveExMemberDetails_(member){// still needs refining...
  var balance = member.getCurrentBalance()
  var balDate = member.getCurrentBalanceDate()
  var sheet = SpreadsheetApp.getActive().getSheetByName('Ex Members')
  if (isDRY){
    sheet.appendRow(['',member.name, member.id, balance, '', '', balDate,
                                                                     '', '', '', '', '', '', 
                                                                     member.id, member.name, member.email, member.mobile , member.homePhone])
  } else {
    SpreadsheetApp.getActive().getSheetByName('Ex Members').appendRow(['',member.name, member.id,  '', '', balDate, balance,
                                                                     50, '', '', '', '', '', 
                                                                     member.id, member.name, member.email, member.mobile , member.homePhone, member.address]) 
    
  }
  var row = sheet.getLastRow()
  sheet.getRange(row,  9).setFormulaR1C1('=sumifs(bank_Amount, bank_Bin, R[0]C[-6], bank_Date, ">" & R[0]C[-3], bank_Amount, ">0")')
  sheet.getRange(row, 10).setFormulaR1C1('=sumifs(bank_Amount, bank_Bin, R[0]C[-7], bank_Date, ">" & R[0]C[-4], bank_Amount, "<0")')
  sheet.getRange(row, 12).setFormulaR1C1('sum(R[0]C[-5]:R[0]C[-1])')

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
    log('Failed to remove member - member not found on Orders sheet', member.id, member.name)
    return
  }
  var col = ids.indexOf(member.id)
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
  var col = ids.indexOf(member.id)
  ss.getSheetByName('Totals').deleteColumn(col+1)
  log(['Removed member from Totals', member.id, member.name])
}

//---------------------------------------------------
// this code needs to go to CoopLib when ready

function removeFromCurrentContacts(member) {
  //  Remove contact from current list - has to be run by coop account
  if (isFRESH){
    var coopGroup = ContactsApp.getContactGroup("Co-op members")  
    var exGroup = ContactsApp.getContactGroup("Ex members")
    var contacts = ContactsApp.getContactsByName(member.name)
    if (contacts.length == 0) {
      log(["Contact not found", member.name])
    } else {
      for (var i in contacts) {
        exGroup.addContact(contacts[i])
        coopGroup.removeContact(contacts[i])
        log(["Moved contact from Co-op Members group to Ex Members group", contact[i].name])
      }
    }
  }
}
