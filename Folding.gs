// Folding

// 1.0 Code created to assist with winding up FRESH Jan 2020 - sending messages with balances to members 
 
function closeLastPaymentDate(id) {
  var member = getMember(id)
  return member.getLatestPayment().date
}

function tt (){
  
}

function notifyThisMemberOfBalance(){  // created to suit winding up of co-op - needs work to go live with Dry
  // call from Totals sheet or from Ex Members sheet to notify member of current balance
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
//      var response = ui.alert("Remove " +  sheet.getRange(TOT_ID_ROW-1, thisCol).getValue() + " from the co-op?", ui.ButtonSet.YES_NO)
//      if (response == ui.Button.YES) {
     // notifyBalance(//sheet.getRange(TOT_ID_ROW, thisCol).getValue())
//    }
    }
    return
  } 
  else if (sheetName == 'Ex Members'){
    if (thisRow < 45){
      ui.alert("Please move to a member row. (Not tested  on rows 4 to 44).")
      return
    }
    var member = new Member()
    var data = sheet.getRange(thisRow, 1, 1, sheet.getLastColumn()).getValues()[0]
    member.name = data[1].toString().trim()
    var matches = /(\w+[-']?\w*)\W+(.*)/.exec(member.name)
    member.firstName = matches[1]
    member.lastName = matches[2]
    
    member.id = data[2].toString()
    member.removalDate = new Date(data[5])
    member.closingBalance = data[6]
    member.netBalance = data[11]
    member.email = data[16]
    notifyBalance(member)
  } 
  else {ui.alert("Select a cell in the member's row in Ex Members, or in the member's column in Totals and try again.")}  
}


function notifyBalance(member, optUrl){
  var subject = member.getFullName() + " has left the " + (isFRESH ? "Fresh" : "Dry") + " co-op"
  member.Balance = Number(member.getCurrentBalance()) + MEMBERSHIP_BOND
  var details = formatFinalAcctDetails_(member)
  
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
    var closure = "At the recent AGM of Kapiti Fresh co-op, members reluctantly voted to wind up the co-op." + brbr
    
    // notify Member, Treasurer and IT
    var action = (member.netBalance > 0 
      ? " Please forward your account details to " + TREASURER_EMAIL + " so that " + TREASURER_NAME + " can arrange a refund." + brbr
      : "Please contact our treasurer "
         + TREASURER_NAME + " at " + TREASURER_EMAIL + " if you wish to make special payment arrangements." + brbr);
  
    MailApp.sendEmail({
      to:[IT_EMAIL].join(','),    //member.email, IT_EMAIL, TREASURER_EMAIL
      subject: "Fresh co-op credit - " + member.name,
      htmlBody: "Hi " + member.firstName + brbr 
      + closure
      + "Your " 
      + (isFRESH ? "Fresh" : "Dry") 
      + " co-op account has been deleted. Your net balance is $" 
      + Math.abs(member.netBalance).toFixed(2) 
      + (member.netBalance<0 ? " in debit." : " in credit.")
      + brbr
      + action
      + details
      + recentPayments
      + autoGenMsg
    })
  }
}
   
function formatFinalAcctDetails_(member){
  var date = Utilities.formatDate(member.removalDate, "GMT+12:00", "d MMMM yyyy")
  var details = "<table>"
  details += "<tr><td>ID</td><td>" + member.id + "</td></tr>"
  details += "<tr><td>Name</td><td>" + member.getFullName() + "</td></tr>"
  details += "<tr><td>Balance Date</td><td>" + date + "</td></tr>"
  
  details += (isFRESH ? "<tr><td>Closing Balance</td><td>" + "$" + Math.abs(member.closingBalance).toFixed(2)
                                                         + (member.closingBalance>=0 ? " in credit" 
                                                                                          : " in debit")

                        + "</td></tr>"
                        + "<tr><td>Bond Refund</td><td>$50.00</td></tr>" 
                      : "")
  details += "<tr><td>Net Balance</td><td>" + "$" + Math.abs(member.netBalance).toFixed(2)
                                                  +(member.netBalance>=0 ? " in credit" 
                                                                           : " in debit")
                                                  + "</td></tr>"
  details += "</table>"       

  return details
} 