// REMINDERS
// v0.7

// merged/synched 4-Mar-18
 
// calls CoopLib, which can be found at https://script.google.com
// go to CoopLib to update phone details when they change
//
// editing needed here to generalise call - say sendText(msg, members)

function sendReminderSMS(){
  var members = haventOrdered()
  
//  var days = ['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'];
//  var day = days[ Now.getDay() ];

//  var message = "A reminder to Fresh co-op members who have not yet ordered:\n" +
//         "FRESH co-op orders will be capped at 10:00 this morning. \n" +
//         "You may order up to a complete bin until 8:00 this evening.\n" +
//                  "https://docs.google.com/spreadsheets/d/1TIvspN97J305bRWz6Ob9aHozu1eQ_Pp3HOf-O6byp1k/edit#gid=18 "
//var message = "FRESH Co-op orders are open unil 10am MONDAY"
  
    var message = "A reminder to Fresh co-op members who have not yet ordered:\n  FRESH Orders will close TODAY at " + CLOSE_TIME + 
    ".  "          //\n Please take a moment to order - the co-op works best when most members order. "  

  var mobile
  var re = /\(*02\d/;

  Logger.log(JSON.stringify(members, null, 4)) 
  
  for (var i=0; i < members.length; i++){
    mobile = getMobile(members[i][0])
    if (re.test(mobile)) {
      CoopLib.sendSMS(mobile, message)
      Logger.log(mobile)
    }
  }
}

function sendText(msg, numbers){
  
}

function testMobile(){
  //Logger.log(getMobile("1059"))
  CoopLib.sendSMS("0211653763", "FRESH co-op orders will be capped at 10am. \n" +
                  "You may order up to a complete bin until 8pm.\n" +
                  "https://docs.google.com/spreadsheets/d/1TIvspN97J305bRWz6Ob9aHozu1eQ_Pp3HOf-O6byp1k/edit#gid=18 ")}
                       
function getMobile(id){
  var sheet = SpreadsheetApp.getActive().getSheetByName("Members") 
  var data = sheet.getDataRange().getValues()
  //var name = sheet.getName()
  var index = ArrayLib.indexOf(data, 1, id)
  return index && data[index][5] || ""
  //ArrayLib.indexOf(data, columnIndex, value)
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
      ordered.push([ids[i], members[i], -currOrders[i], provBalances[i]])
    }
  }
  return ordered //.sort(function(a, b){return a[2]>b[2]})
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

function haventOrdered(){
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var currOrders = ss.getRangeByName("tot_Current_Orders").getValues()[0]
  var prevOrders = ss.getRangeByName("tot_Previous_Orders").getValues()[0]
  var provBalances = ss.getRangeByName("tot_Provisional_Balances").getValues()[0]
  var ids = ss.getRangeByName("tot_IDs").getDisplayValues()[0]
  var names = ss.getRangeByName("tot_members").getValues()[0]
  var ordered = []

  var thisid = ""
  
  for (var i =0; i < currOrders.length; i++){
    thisid = ids[i]

    if (currOrders[i] == -MIN_ORDER_FEE){
     ordered.push([thisid, names[i], currOrders[i]])  //, members.id[thisid].mobile])
    }
  }
  return ordered
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
