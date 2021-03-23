// REMINDERS

/*
Send SMS Messages
=================
sendReminderSMS        - send reminders to everyone who hasn't ordered - called from menu and from timed trigger - calls getMobilesHaventOrderedButDidLastTime to select members
sendAlertHaveOrdered   - send SMS for late delivery, change of pack days etc                                     - calls getMobilesHaveOrdered and sendText 
sendText(msg, mobiles) - sends SMS to each number in list of Mobiles

Email Reminders
===============
emailRemindersHaventOrdered          - 
emailReminder(member)                - 


Select Members (or members data)  - SOME OR MANY NOT WORKING - NEEDS TO BE LOOKED AT TO GET NOTIFICATIONS WORKING *************************************************************
================================
getMobile              - get Mobile from Members sheet (slow)    ...not currently in use
refreshSheet           - cause reminders sheet to recalculate

madeAnOrder            - get (some data of) members who have ordered
getMobilesHaveOrdered  - get mobiles of members who have ordered
getMemberssHaventOrdered              - 
getMembersHaventOrderedButDidLastTime - 
getMembersHaventOrderedAgain          - 
getPackTeam            - reading pack team from roster - for notifications... not in use yet
haventOrderedThisTime  - 
HaventOrderedAgain     - 
haventOrdered          - 
getMobilesHaventOrderedButDidLastTime - 
suspended              - 
*/

// v0.94 Fixed dry sms reminders 22/3/21
// v0.93 Added doc at top, triying to fix the selection routines, which are mostly stuffed!
// v0.92 Send alert to members who have ordered
// v0.91 Edits for generalising changes - under devt - works for Dry not tested Fresh yet - changes to getting members who have not ordered
// v0.9 Adjust for new Dry Members sheet layout
// v0.8 Make it work for Dry
// v0.7f
// merged/synched 4-Mar-18
 
// CoopLib.sendSMS can be found at https://script.google.com or just using File/Open
// go to CoopLib to update phone details when they change
//


//=======================================================================
// Send SMS Reminders
//=======================================================================
  
function testMe1(){
//  var numbers = [getMobile('7177')]
//  sendText("from testMe", numbers)
  var x = getMobilesHaventOrdered()
  log (x.length)
}

function sendReminderSMS(){
  var weekdays = ['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'];
  var todayi = new Date().getDay()
  var today = weekdays[todayi]
  var tomorrow = weekdays[todayi+1%7]
  var closeDay
  
  if (CLOSE_DAY == today){// if it is closing day today we want to say "closes today" not "closes Monday" (or whatever day) 
    closeDay = "today"
  }
  else if (CLOSE_DAY == tomorrow) {
    (closeDay = "tomorrow") 
  } else {closeDay = CLOSE_DAY}
  
  var re = /\(*02\d/;

  if (isFRESH) {
    var message = "A reminder to Fresh co-op members who have not yet ordered:\n  FRESH Orders will close " + closeDay +" at " + CLOSE_TIME + ".  "          //\n Please take a moment to order - the co-op works best when most members order. "  
    var members = haventOrdered()  
    var mobile
        
    for (var i=0; i < members.length; i++){
      mobile = getMobile(members[i][0])
      if (re.test(mobile)) {
        CoopLib.sendSMS(mobile, message)
        log(["Sent reminder to ", mobile])
      }
    }
  } 
  else {
    var message = "A reminder to DRY co-op members who have not yet ordered:\n Dry orders will close " + closeDay +" at " +       CLOSE_TIME + ".  "
    var mobiles = getMobilesHaventOrderedButDidLastTime()
    for (var i=0; i < mobiles.length; i++){
      if (re.test(mobiles[i])) {
        CoopLib.sendSMS(mobiles[i], message)
        log(["Sent reminder to ", mobiles[i]])
      }
    }
  }
}

function sendAlertHaveOrdered(){
  var message = "Dry goods pick-up has been postponed until Tuesday and Wednesday because the food has not arrived.\n\n Please let Nico know if you can help unpack the pallet on Monday."
  var mobiles = getMobilesHaveOrdered()
  sendText(message, mobiles)
}

function sendText(message, mobiles){
  var re = /\(*02\d/;
  for (var i=0; i < mobiles.length; i++){
    if (re.test(mobiles[i])) {
      CoopLib.sendSMS(mobiles[i], message)
      log(["Sent reminder to ", mobiles[i]])
    }
  }
}

                     
function getMobile(id){
  var sheet = SpreadsheetApp.getActive().getSheetByName("Members") 
  var data = sheet.getDataRange().getValues()
  //var name = sheet.getName()
  var index = ArrayLib.indexOf(data, MEM_ID_OFFSET, id)
  return (index >= 0) && data[index][MEM_MOBILE_OFFSET] || ""
}


//=======================================================================
// Send Email Reminders
//=======================================================================


function emailRemindersHaventOrdered(){
  var members = getMembersHaventOrderedAgain().sort(dynamicSort('id')).slice(19,30)
  //members = [getMember(7000)] // to test one member
  
  if (members) {members.map(member => emailReminder(member))}
}
  
function emailReminder(member){
  var autoGenMsg = brbr + "<small>This message was automatically generated. Please contact " +
                    IT_NAME + " at " +  IT_EMAIL + " if you have any queries."; 
  var recentPayments = formatPayments_(member)
  var balance = member.getCurrentBalance()
  
  var message =  
         ["Hi ", member.firstName,
          brbr, "Dry co-op orders are still open. Don't forget to place your order by Monday at 6 pm.",
          brbr, "On ", 
          Utilities.formatDate(member.getCurrentBalanceDate(), "GMT+12:00", "d MMMM yyyy"), 
          " your account was $", Math.abs(balance).toFixed(2),
          (balance<0 ? " in debit." : " in credit."),
          recentPayments,
          autoGenMsg].join("")
  
  log (["Emailing...", member.id, member.name])
  notify([member.email, IT_EMAIL], "DRY co-op reminder " + member.id, message)  //member.email
  //sendText(message, ["0211653763"]) //not html
          
}


//=======================================================================

function refreshSheet(){
  SpreadsheetApp.getActiveSheet().insertRowBefore(1).deleteRow(1)
}


function madeAnOrder(){// returns data in arrays suitable for sheet
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var currOrders = ss.getRangeByName("tot_Order_Subtotals").getValues()[0]

  var provBalances = ss.getRangeByName("tot_Provisional_Balances").getValues()[0]
  var ids = ss.getRangeByName("tot_IDs").getValues()[0]
  var memberList = ss.getRangeByName("tot_members").getValues()[0]
  var ordered = []
  
  for (var i =0; i < currOrders.length; i++){
    if (currOrders[i] >0) { // < -MIN_ORDER_FEE){
      ordered.push([ids[i], memberList[i], currOrders[i], rounded(provBalances[i])])
    }
  }
  return ( ordered) //.sort(function(a, b){return a[2]>b[2]})

}



function getMobilesHaveOrdered(){//Dry

  var ss = SpreadsheetApp.getActiveSpreadsheet()

  var totIDs = ss.getRangeByName("tot_IDs").getDisplayValues()[0]
  var currOrders = ss.getRangeByName("tot_Current_Orders").getValues()[0]
  var names = ss.getRangeByName("tot_members").getValues()[0]
  
  var memIDs = ArrayLib.transpose(ss.getRangeByName("mem_IDs").getValues())[0]
  var memMobiles = ArrayLib.transpose(ss.getRangeByName("mem_Mobiles").getValues())[0]
  memIDs.shift()      // remove header
  memMobiles.shift()  // remove header
  
  var mobiles = []
  var totIndex
   
  // step through members list, check if they ordered, add mobile to list
  
  for (var i =0; i < memIDs.length; i++){
    totIndex = totIDs.indexOf(memIDs[i])
    if ((totIndex >= 0) && (currOrders[totIndex] < -MIN_ORDER_FEE) && memMobiles[i] != "") {
      mobiles.push(memMobiles[i])
    }
  }
  
  return mobiles
}                  
 


function getMembersHaventOrdered(){
  return members = getMembers().filter(x => MIN_ORDER_FEE &&                           // check they're not suspended
                                            x.getCurrentOrder() == -MIN_ORDER_FEE)
}

function getMembersHaventOrderedButDidLastTime(){
  return members = getMembers().filter(x => MIN_ORDER_FEE &&
                                            x.getCurrentOrder() == -MIN_ORDER_FEE && 
                                            x.getPreviousOrder() < -MIN_ORDER_FEE)
}
                    
function getMembersHaventOrderedAgain(){
  return members = getMembers().filter(x => MIN_ORDER_FEE &&
                                            x.getCurrentOrder() == -MIN_ORDER_FEE &&
                                            x.getPreviousOrder() == -MIN_ORDER_FEE)
}

function getPackTeam(){
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var teamList = ss.getRangeByName("Roster_This_Pack").getValues().slice(2)
  var team
  
  for (var i =0; i < teamList.length; i++){
    if (teamList[i].length>0){
      team.push([team[i].slice(0,3)])
    }
  }
  return team
}


function haventOrderedThisTime(){
  var members = getMembersHaventOrderedButDidLastTime()
  return members.map(m => [m.id, m.name, ,m.mobile])
}

function HaventOrderedAgain() {
  var members = getMembersHaventOrderedAgain()
  return members.map(m => [m.id, m.name, ,m.mobile])
}
  
//  var ss = SpreadsheetApp.getActiveSpreadsheet()
//  //var totals = ss.getSheetByName("Totals").getDataRange().getValues()
//  var currOrders = ss.getRangeByName("tot_Current_Orders").getValues()[0]
//  var prevOrders = ss.getRangeByName("tot_Previous_Orders").getValues()[0]
//  var provBalances = ss.getRangeByName("tot_Provisional_Balances").getValues()[0]
////  var currBalances = ss.getRangeByName("tot_Current_Balances").getValues()[0]
//  var ids = ss.getRangeByName("tot_IDs").getDisplayValues()[0]
//  var names = ss.getRangeByName("tot_members").getValues()[0]
//  var ordered = []
//
//  var thisid = ""
//  
//  for (var i =0; i < currOrders.length; i++){
//    thisid = ids[i]
//
//    if (currOrders[i] == -MIN_ORDER_FEE && prevOrders[i] <-MIN_ORDER_FEE){// modified for fresh
//     ordered.push([thisid, names[i], currOrders[i]])  //, members.id[thisid].mobile])
//    }
//  }
//  return ordered


function haventOrdered(){// excludes members on leave
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var currOrders = ss.getRangeByName("tot_Current_Orders").getValues()[0]
  var prevOrders = ss.getRangeByName("tot_Previous_Orders").getValues()[0]
  var provBalances = ss.getRangeByName("tot_Provisional_Balances").getValues()[0]
  var members = getMembers()
  var ids = ss.getRangeByName("tot_IDs").getDisplayValues()[0]
  var names = ss.getRangeByName("tot_members").getValues()[0]
  var ordered = []

  var thisid = ""
  
  for (var i =0; i < currOrders.length; i++){
    thisid = ids[i]

    if (currOrders[i] == -MIN_ORDER_FEE){
      if (members[thisid]) {ordered.push([thisid, names[i] ])}      ///, getMobile(thisid)])  //, members.id[thisid].mobile])
    }
  }
  return ordered
}

function getMobilesHaventOrderedButDidLastTime(){// excludes members on leave  //only tested on Dry
  var members = getMembersHaventOrderedButDidLastTime()
  return ( members.map(m => m.mobile))
    
//  var ss = SpreadsheetApp.getActiveSpreadsheet()
//
//  var totIDs = ss.getRangeByName("tot_IDs").getDisplayValues()[0]
//  var currOrders = ss.getRangeByName("tot_Current_Orders").getValues()[0]
//  var names = ss.getRangeByName("tot_members").getValues()[0]
//  
//  var memIDs = ArrayLib.transpose(ss.getRangeByName("mem_IDs").getValues())[0]
//  var memMobiles = ArrayLib.transpose(ss.getRangeByName("mem_Mobiles").getValues())[0]
//  memIDs.shift()      // remove header
//  memMobiles.shift()  // remove header
//  
//  var mobiles = []
//  var totIndex
//   
//  // step through members list, check if they ordered, add mobile to list
//  
//  for (var i =0; i < memIDs.length; i++){
//    totIndex = totIDs.indexOf(memIDs[i])
//    if ((totIndex >= 0) && (currOrders[totIndex] == -MIN_ORDER_FEE) && memMobiles[i] != "") {
//      mobiles.push(memMobiles[i])
//    }
//  }
//  
//  return mobiles
}




                   
//function haventOrderedAgain(){
//  var ss = SpreadsheetApp.getActiveSpreadsheet()
//  //var totals = ss.getSheetByName("Totals").getDataRange().getValues()
//  var currOrders = ss.getRangeByName("tot_Current_Orders").getValues()[0]
//  var prevOrders = ss.getRangeByName("tot_Previous_Orders").getValues()[0]
////  var currBalances = ss.getRangeByName("tot_Current_Balances").getValues()[0]
//  var provBalances = ss.getRangeByName("tot_Provisional_Balances").getValues()[0]
//  var ids = ss.getRangeByName("tot_IDs").getDisplayValues()[0]
//  var names = ss.getRangeByName("tot_members").getValues()[0]
//
//  var thisid = ""
//  var list = []
//  
//  for (var i =0; i < currOrders.length; i++){
//    thisid = ids[i]
//
//    if (currOrders[i] == -MIN_ORDER_FEE && prevOrders[i] >= -MIN_ORDER_FEE){// didn't order last time either ... modified for fresh
//      list.push([thisid, names[i], currOrders[i]])  //, members.id[thisid].mobile])
//    }
//  }
//  return (list.length>0 && list) || [[,"No members"]]
//}


function suspended(){
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  //var totals = ss.getSheetByName("Totals").getDataRange().getValues()
  var currOrders = ss.getRangeByName("tot_Current_Orders").getValues()[0]
  var prevOrders = ss.getRangeByName("tot_Previous_Orders").getValues()[0]
  //var currBalances = ss.getRangeByName("tot_Current_Balances").getValues()[0]
  var provBalances = ss.getRangeByName("tot_Provisional_Balances").getValues()[0]
  var ids = ss.getRangeByName("tot_IDs").getDisplayValues()[0]
  var names = ss.getRangeByName("tot_members").getValues()[0]

  var thisid = ""
  var list = []
  
  for (var i =0; i < currOrders.length; i++){
    thisid = ids[i]

    if (currOrders[i] > -MIN_ORDER_FEE){// no fees - suspended, on leave, leaving, left
      list.push([thisid, names[i], currOrders[i], provBalances[i]])  //, members.id[thisid].mobile])
    }
  }
  return list
}



//function refreshFreshDirect(){//don't need this anymore
//  var ss = SpreadsheetApp.getActiveSpreadsheet()
//  var sheet = ss.getSheetByName("T O")
//  sheet.insertRowBefore(1).deleteRow(1)
//}
