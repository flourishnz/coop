//// STATEMENTS
// v0.3  Move email addresses to globals
//// v0.2
//
//// Developing May-June 2018
//
////function showDialog(data) {
////  var template = HtmlService.createTemplateFromFile('testDialog')
////      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
////      .setWidth(400)
////      .setHeight(300);
////      .data = data;
////  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
////      .showModalDialog(html, 'My custom dialog');
////}
//
////...still to do 
//// move contacts code to CoopCoopLib
//// updateEmail
//// updateID ?
//
//
//function testStatements(){
//  say(getTransactions("8078"))
////  var id = "3102"
////  var member = getLatestTotals(id)
////  say (member)
////  say(getBanking(member.id))
////  say (getEmail(member.id))
//  
////  var rtn = getFreshContacts()
////  var data = "Got " + rtn.length + " contacts"
////  showDialog(data)
//}
//
//
////function getTransactions(optId) {// cached to allow for repeated calls over a couple of minutes
////  var id = optId || ""
////  
////  // get transactions from Banking or from Cache
////  var cache = CacheService.getScriptCache();
////  var cached = cache.get("transactions");
////  if (cached == null) {
////    var ss = SpreadsheetApp.getActiveSpreadsheet()
////    var data = ss.getSheetByName("Banking").getDataRange().getValues()
////    var idCol = ss.getRangeByName("bank_Bin").getColumn()-1
////    var amountCol = ss.getRangeByName("bank_Amount").getColumn()-1
////    transactions = []
////    for (var i = data.length-1; i > 0; i--) {
////      transactions.push({date: data[i][0],
////                         amount: data[i][amountCol]
////                        })
////    }
////    cache.put("transactions",  JSON.stringify(transactions))
////  } else {
////    transactions = JSON.parse(cached)
////  }
////  
////  if (isValidId(id)) {
////    return _.filter(transactions, function(x) {return x.id == id}) 
////  } else {
////    return transactions
////  }
////}
//
//function getMemberTransactions(id) {
//  return transactions[id]
//}
//
//
//function getLatestTransactions() {// saves to ScriptProperties
//  // if transactions have already been cached...
//  var cached = CacheService.getScriptCache().get('transactions')
//  if (cached !== null) {return JSON.parse(cached)}
//
//  // otherwise generate from sheet
//  var ss = SpreadsheetApp.getActiveSpreadsheet()
//  var data = ss.getSheetByName("Banking").getDataRange().getValues()
//  var idCol = ss.getRangeByName("bank_Bin").getColumn()-1
//  var amountCol = ss.getRangeByName("bank_Amount").getColumn()-1
//  var transactions = []
//
//  for (var i = 0; i < data.length; i++) {
//    var id = data[i][idCol]
//    if (!isValidId(id)) {continue}
//    
//    if (!(id in transactions)) {transactions[id] = []}
//    transactions[id].unshift({date: data[i][0], 
//                              amount: data[i][amountCol]
//                             })
//    if (transactions[id].length == 4) transactions[id].pop()
//  }
//
//  // then add to cache and return
//  CacheService.getScriptCache().put("transactions",  JSON.stringify(transactions))
//  return transactions
//}
//
//
//function sendStatement(member){
//  var ss = SpreadsheetApp.getActiveSpreadsheet()
//  var endNotes = "<br><br>Balance quoted is as shown on the Totals tab of the latest spreadsheet at the time of sending. " + 
//                 "Recent payments may not yet have shown up in our account." 
//  var html =  
//    '<body>' + 
//      '<h3 style="color: green;">Kapiti Co-op - Statement</h3>' +
//      '<h2> Statement </h2><br />' +
//        '<p> Greetings Earthling </p>' +
//          endNotes +
//    '</body>'
//
//  var message = {to: "flourish.nz@gmail.com"     ,//   + ", " + recipients,
//                 subject: "Fresh TEST statement " + ss.getName(),
//                 htmlBody: html + "<br><br><a href='" + ss.getUrl() + "'>" + ss.getName() + "</a>"
//                }
//  
//  MailApp.sendEmail(message)
//}
//
////  tellIT("<h1>Statement</h1><br>Your Fresh account is ((member.ProvisionalBalance < 0) ? "in debit." : "in credit.") +
////            " Your balance is $" + member.ProvisionalBalance  + "." +
////            
////            "<br>" +
////            //"Your most recent orders were ... <br><br>" +
////            "Your most recent payments were ... <br><br>"  +
////            "Payments made after ... have not yet been downloaded.<br> <br>" +
////            "<div foregroundcolor grey>This message has been automatically generated and is still experimental. </div>" +
////            
////            '<iframe id="forum_embed" ' +
////            'src="javascript:void(0)"' +
////            'scrolling="no"' +
////            'frameborder="0"' +
////            'width="900"' +
////            'height="700">' +
////            '</iframe>' +
////            '<script type="text/javascript">' +
////            "document.getElementById('forum_embed').src =" +
////    'https://groups.google.com/forum/embed/?place=forum/fresh-committee ' +
////   '&showsearch=true&showpopout=true&showtabs=false  '+
////   '&parenturl= + encodeURIComponent(window.location.href)' +
////  "</script>"
//            
//       
//
////function getActiveMembersNegative(){ 
////  var data = getLatestTotals()
////  return _.filter(data, function(m) {return true}) // && m.ProvisionalBalance <0)})
////}
////
////function getActiveMembersPositive(){ 
////  var data = getLatestTotals()
////  return _.filter(data, function(m) {return (m.Contingency > 0 && m.ProvisionalBalance >= 0)})
////}
//
//function getLatestTotals(optID){
//  var ID = optID || ""
//  var ssObj = getLatestSS("Fresh Orders 20")  
//  var totals = getTotals(ssObj)
//
//  if (isValidId(ID)) {     
//    return totals[ID]
//  }
//  else {
//    return totals
//  }
//}
//
//
//function getTotals(ssObj){//tested for Fresh only
//
//// if (!ssObj){return};
//  var ss = SpreadsheetApp.open(ssObj)
//  var arr = []
//
//  var sheet = ss.getSheetByName("Totals")
//  var data = sheet.getDataRange().getValues()
//  
//// Locate, select and transpose the totals section
//  var iSubtotal = ArrayLib.indexOf(data, 0, "Sub-total")
//  var balances = data.slice(iSubtotal, iSubtotal+20)
//  var transBalances = ArrayLib.transpose(balances)
//      
//// Locate headers
//  var iContingency = ArrayLib.indexOf(balances, 0, "Contingency (wastage, errors, admin...)")
//  var iAccounting = ArrayLib.indexOf(balances, 0, "Accounting and IT")
//  var iMembership = ArrayLib.indexOf(balances, 0, "Membership")
//  var iOrderTotal = ArrayLib.indexOf(balances, 0, "Order Total")
//  
//  var iNames = ArrayLib.indexOf(balances, 0, "Fresh Accounts")
//  var iIDs = ArrayLib.indexOf(balances, 0, "Totals")
//  
//  var iPreviousBalance = ArrayLib.indexOf(balances, 0, "Previous Order")
//  var iPreviousOrder = ArrayLib.indexOf(balances, 4, "Previous Order")
//  var iPreviousPayments = ArrayLib.indexOf(balances, 4, "payments received between...")
//  var iPreviousOtherCredits = ArrayLib.indexOf(balances, 4, "Other credits and debits")
//  
//  var iCurrentBalance = ArrayLib.indexOf(balances, 0, "Current Order")
//  var iCurrentOrder = ArrayLib.indexOf(balances, 4, "Current order total (as above)")
//  var iCurrentPayments = iCurrentOrder + 1
//
//
//  var iProvisionalBalance = ArrayLib.indexOf(balances, 4, "Provisional Balance")
//  
//// initalise 
//  var member = []
//  var id = ""
//  var previousBalanceDate = new Date(balances[iPreviousBalance, 4])
//  var currentBalanceDate = new Date(balances[iCurrentBalance, 4])
//  var provisionalBalanceDate = new Date(balances[iCurrentPayments + 1, 4])
//  var totals = {}
//
//// step through the members
//  
//  for(var i = 5; i < balances[6].length; i++){   // 5 is the first member column - get this from somewhere?
//
//    member = transBalances[i];
//    id = member[iIDs].toString()
//    
//    if (isValidId(id)) {
//      totals = {
//        ID: id,
//        Name: member[iNames],
//        
//        Contingency: member[iContingency],
//        Accounting: member[iAccounting],
//        Membership: member[iMembership],
//        OrderTotal: member[iOrderTotal],
//        
//        PreviousBalance: member[iPreviousBalance],
//        PreviousOrder: member[iPreviousOrder],
//        PreviousPayments: member[iPreviousPayments],
//        PreviousOtherCredits: member[iPreviousOtherCredits],
//        
//        CurrentBalance: member[iCurrentBalance],
//        CurrentOrder: member[iCurrentOrder],
//        CurrentPayments: member[iCurrentPayments],
//        
//        ProvisionalBalance: member[iProvisionalBalance]
//      }
//
//      arr[id] = totals
//    }
//  }
//  return arr
//}
//
//
//
//function isValidId(id) {
//  return /^\d{4}$/.test(id)
//}
//
//
///*--------------------------------------------------------
//getMembers() - [objects] from Members tab
//
//adddMemberToContacts(member) - if member not in contacts addContact
//                               else if member in once then updateContact
//                               else fail
//                               
//addContact(member, group)
//updateContact(member, group)
//
//                               
//                          
//-----------------------------------------------------------*/
