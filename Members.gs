// MEMBERS
// v1.4

// 15-3-18 fixed bug ss not defined in function Member and SYNCHED
// 1-3-18 synched - remove member
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
    var ids = ArrayLib.transpose(ss.getRangeByName("tot_Bins").getValues())
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
  addMemberToContacts(member)
  
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


//// MEMBERS
//
//
function getMembers(){// array of objects
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


function getMember(id){
  var sheet = SpreadsheetApp.getActive().getSheetByName("Members") 
  var data = sheet.getDataRange().getValues()
  var member = new Member();
 
  var i = ArrayLib.indexOf(data, MEM_ID_OFFSET, id)   // look for id
  if (!i)  {return }
  
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
  Logger.log(member.getCurrentBalanceDate())
  Logger.log(member.getCurrentBalance())
  return member
}


function testm(){
  removeMember("8175")
}

