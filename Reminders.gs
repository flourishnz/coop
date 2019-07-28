// REMINDERS
// v0.92 Send alert to members who have ordered
// v0.91 Edits for generalising changes - under devt - works for Dry not tested Fresh yet - changes to getting members who have not ordered
// v0.9 Adjust for new Dry Members sheet layout
// v0.8 Make it work for Dry
// v0.7f

// merged/synched 4-Mar-18
 
// calls CoopLib, which can be found at https://script.google.com
// go to CoopLib to update phone details when they change
//
// editing needed here to generalise call - say sendText(msg, members)

//Sally 8080
//Valamaya 8146
//Erica 7213
//Victoria Garlick 8069

function test(){
  var mobile = "0211653763"
  var message = "Test no spacing"
  sendText("sendText " + message, mobile)
  CoopLib.sendSMS(mobile, "sendSMS " + message)
}
  
function testMe(){
  say (getMobile('7177'))
  var numbers = [getMobile('7177')]
  sendText("from testMe", numbers)
}

function sendReminderSMS(){// haven't tried to re-merge these two methods... 
  // methods were split to allow for new dry - too many ex-members on sheet after merge - only want to notify ones who are current (filled in form)
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
    
  if (isFRESH) {   
    var message = "A reminder to Fresh co-op members who have not yet ordered:\n  FRESH Orders will close " + closeDay +" at " + CLOSE_TIME + ".  "          //\n Please take a moment to order - the co-op works best when most members order. "  
  } else {
    var message = "A reminder to DRY co-op members who have not yet ordered:\n Dry orders will close " + closeDay +" at " + CLOSE_TIME + ".  "
  }
  
  var re = /\(*02\d/;

  if (isFRESH) {
    var members = haventOrdered()  
    var mobile
    
    Logger.log(JSON.stringify(members, null, 4)) 
    
    for (var i=0; i < members.length; i++){
      mobile = getMobile(members[i][0])
      if (re.test(mobile)) {
        CoopLib.sendSMS(mobile, message)
        log(["Sent reminder to ", mobile])
      }
    }
  } 
  else {
    var mobiles = getMobilesHaventOrdered()
    for (var i=0; i < mobiles.length; i++){
      if (re.test(mobiles[i])) {
//        CoopLib.sendSMS(mobiles[i], message)
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

function refreshSheet(){
  SpreadsheetApp.getActiveSheet().insertRowBefore(1).deleteRow(1)
}

function refreshFreshDirect(){
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var sheet = ss.getSheetByName("T O")
  sheet.insertRowBefore(1).deleteRow(1)
}

function madeAnOrder(){
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  //var totals = ss.getSheetByName("Totals").getDataRange().getValues()
  var currOrders = ss.getRangeByName("tot_Current_Orders").getValues()[0]
  var provBalances = ss.getRangeByName("tot_Provisional_Balances").getValues()[0]
  var ids = ss.getRangeByName("tot_IDs").getValues()[0]
  var members = ss.getRangeByName("tot_members").getValues()[0]
  var ordered = []
  
  for (var i =0; i < currOrders.length; i++){
    if (currOrders[i] < -MIN_ORDER_FEE){
      ordered.push([ids[i], members[i], -currOrders[i], rounded(provBalances[i])])
    }
  }
  return ordered //.sort(function(a, b){return a[2]>b[2]})
}



function getMobilesHaveOrdered(){//only tested on Dry
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


function haventOrderedButDidLAstTime(){
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  //var totals = ss.getSheetByName("Totals").getDataRange().getValues()
  var currOrders = ss.getRangeByName("tot_Current_Orders").getValues()[0]
  var prevOrders = ss.getRangeByName("tot_Previous_Orders").getValues()[0]
  var provBalances = ss.getRangeByName("tot_Provisional_Balances").getValues()[0]
//  var currBalances = ss.getRangeByName("tot_Current_Balances").getValues()[0]
  var ids = ss.getRangeByName("tot_IDs").getDisplayValues()[0]
  var names = ss.getRangeByName("tot_members").getValues()[0]
  var ordered = []

  var thisid = ""
  
  for (var i =0; i < currOrders.length; i++){
    thisid = ids[i]

    if (currOrders[i] == -MIN_ORDER_FEE && prevOrders[i] <-MIN_ORDER_FEE){// modified for fresh
     ordered.push([thisid, names[i], currOrders[i]])  //, members.id[thisid].mobile])
    }
  }
  return ordered
}

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

function getMobilesHaventOrdered(){// excludes members on leave  //only tested on Dry
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
    if ((totIndex >= 0) && (currOrders[totIndex] == -MIN_ORDER_FEE) && memMobiles[i] != "") {
      mobiles.push(memMobiles[i])
    }
  }
  
  return mobiles
}




                   
function haventOrderedAgain(){
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  //var totals = ss.getSheetByName("Totals").getDataRange().getValues()
  var currOrders = ss.getRangeByName("tot_Current_Orders").getValues()[0]
  var prevOrders = ss.getRangeByName("tot_Previous_Orders").getValues()[0]
//  var currBalances = ss.getRangeByName("tot_Current_Balances").getValues()[0]
  var provBalances = ss.getRangeByName("tot_Provisional_Balances").getValues()[0]
  var ids = ss.getRangeByName("tot_IDs").getDisplayValues()[0]
  var names = ss.getRangeByName("tot_members").getValues()[0]

  var thisid = ""
  var list = []
  
  for (var i =0; i < currOrders.length; i++){
    thisid = ids[i]

    if (currOrders[i] == -MIN_ORDER_FEE && prevOrders[i] >= -MIN_ORDER_FEE){// didn't order last time either ... modified for fresh
      list.push([thisid, names[i], currOrders[i]])  //, members.id[thisid].mobile])
    }
  }
  return (list.length>0 && list) || [[,"No members"]]
  
}


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
