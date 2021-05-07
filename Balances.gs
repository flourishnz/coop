/* BALANCES

 v0.1

 Getting all past balances for a particular customer
 --------------------------------------------------
 getCustomerHistory() - NOT IN USE? - goes back through many ss, applying getMemberAccounts to each.  Kinda works but super slow - ok for one member, not for everyone
                                    - gets payments but not returning them? see also Statements:getTotals
 getMemberAccounts(id, ss) - gets ALL payment information (unlimited) and the balances from SS for one member only
 
 
 Sending out current balances to members
 ---------------------------------------
 ... add code to step through all contacts in sensible order - currently limited to group method because Gmail limits generated emails to 35ish per day on free account
 
 emailGroup(group) - calls sendReleaseEmail for each member of the contact group - currently defaults to "G0 partial"
 sendReleaseEmail(contact) - sends individualised release email to a contact - WORKING 2020
 
 SAMPLEformatReleaseAcctDetails_(member) - NOT CALLED = possibly still the same as the one used when closing accounts
*/

function makeSendGroup(max = 3){ 
  var ids = ["7177"] //getNotSent().slice(0, max)
  var sendGroup = ContactsApp.getContactGroup("sendGroup")
  ids.map(x => addToGroup(x, sendGroup))
}


function getNotSent(){
  // get data
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var idsRange = ss.getRangeByName("tot_IDs")
  var totals = ss.getSheetByName("Totals")
  var balances = totals.getRange(idsRange.getRow(), idsRange.getColumn()+1, 13, idsRange.getNumColumns() ).getDisplayValues()

  // transpose, select not sent, sort to find members who owe the most, pass back just the ids
  return ArrayLib.transpose(balances)
    .filter(x => x[12] !== "sent")
    .sort((a, b) => a[6] - b[6])
    .map(x => x[0])
}


function markSent(id="7177", 
                  ids=SpreadsheetApp.getActive().getRangeByName("tot_IDs").getDisplayValues()){
  ss = SpreadsheetApp.getActive()
  const statusRange = ss.getRangeByName("tot_Status")
  const row = statusRange.getRow()
  const col = ids[0].indexOf(id)
  ss.getSheetByName("Totals").getRange(row, col).setValue("sent")
}

function getCustomerHistory(){  
  const sss = getSsSortByName("^Dry Orders Merged*")
  var accounts = (sss.map(ssfile => getMemberAccounts("8204", SpreadsheetApp.open(ssfile))))
  //var keys = accounts[0].getKeys()

}

function getMemberAccounts(id="7177", ss=SpreadsheetApp.getActiveSpreadsheet()) {
  var member = getMember(id, ss)
  var payments = member.getPayments()
  return {pbd: member.getPreviousBalanceDate(), 
          pb: member.getPreviousBalance(),
          po: member.getPreviousOrder(),
          pc: member.getPreviousCredits(),
          cbd: member.getCurrentBalanceDate(),
          cb: member.getCurrentBalance()}
}



/************************************************************************
Release emails using a contact list -
  get contacts from list eg "g4 members S-Z"
  for each contact
    emailrelease - pass in the preferred email addresses from the contact instead of using the one on the member's record (???)
****************************************************************************/

function emailGroup(group = "G0 partial"){
  const contactsGroup = ContactsApp.getContactGroup(group)
  if (contactsGroup == null) {
    throw("Wrong account? Contacts group not found: "+group)
    return
  }
  
  contactsGroup.getContacts().filter(c => hasID(c))
                             .map(contact => sendReleaseEmail(contact))  
}

function sendReleaseEmail(contact){
  var id = getID(contact)
  var member = getMember(id)
  if (member.length == 0) {
    tellIT([contact.getFullName(), id, "is in the contacts list but not in the members list. Not notified."].join(" "))
    return
  }
  var emails = contact.getEmails()
  var greeting = ["Hi ", contact.getGivenName(),
                  brbr, "Dry co-op orders are currently open. Don't forget to place your order by ", 
                  CLOSE_DAY, " at ", 
                  CLOSE_TIME,
                  " (although Nico does sometimes close the orders later than this)."].join("")
  var prevOrder = member.getPreviousOrder()
  prevOrder == -2 ? prevOrder = ""
                  : prevOrder = brbr + "Your previous order totalled $" + Math.abs(prevOrder).toFixed(2) + "."
  var recentPayments = formatPayments_(member, 3)

  var prevBalance = member.getPreviousBalance()
  var prevBalanceMsg = [brbr, "On ", 
                 Utilities.formatDate(member.getPreviousBalanceDate(), "GMT+12:00", "d MMMM yyyy"), 
                 " your account was $", Math.abs(prevBalance).toFixed(2),
                 (prevBalance<0 ? " in debit." : " in credit.")].join("")
                 
  var currBalance = member.getCurrentBalance()
  var currBalanceMsg = [brbr, "On ", 
                 Utilities.formatDate(member.getCurrentBalanceDate(), "GMT+12:00", "d MMMM yyyy"), 
                 " your account was $", Math.abs(currBalance).toFixed(2),
                 (currBalance<0 ? " in debit." : " in credit.")].join("")
  
  var autoGenMsg = brbr + "<small>This message was automatically generated. Please contact " +
                    IT_NAME + " at " +  IT_EMAIL + " if you have any queries."; 
  
  var message = [greeting,
                 currBalanceMsg,
                 prevOrder,
                 recentPayments,
                 autoGenMsg].join("")
  
        emails.map(e => {notify(e.getAddress(), "DRY orders open "+id,    message)
                         log (["Emailed...", id, contact.getGivenName(), e.getAddress(), "member: "+member.firstName])}
  
  )
  //sendText(message, ["0211653763"]) //not html

}



function SAMPLEformatReleaseAcctDetails_(member){ // use as model if required to improve above emails
  var pbDate = Utilities.formatDate(member.previousBalanceDate, "GMT+12:00", "d MMMM yyyy")
  var cbDate = Utilities.formatDate(member.previousBalanceDate, "GMT+12:00", "d MMMM yyyy")

  var details = "<table>"
  details += "<tr><td>ID</td><td>" + member.id + "</td></tr>"
  details += "<tr><td>Name</td><td>" + member.getFullName() + "</td></tr>"
  //details += "<tr></tr>
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


/************************************************************************
Previousy released emails using list of ids called tempIDs - 
  WORKED but had problem with couples,as sent one message per member listing but primary contact received both
*************************************************************************/

//function temptemp(ss = SpreadsheetApp.getActive()){
//  var range = ss.getRangeByName("tempIDs")
//  var tempIDs = range.getDisplayValues().filter(x => x[1]=="").slice(0,15)  //grab first 15 ids that have not been emailed yet
//  tempIDs.map(x => emailRelease(getMember(x[0])))                                   
//}



