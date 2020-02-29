//Roster report 30-1-20
// not complete - doesn't look right - eg first function doesn't retun anything yet!  29-2-20 regex functions not working either

function getCurrentRoster() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var range = ss.getRangeByName("ros_This_Pack")
  var arr = range.getValues()
  //var last = sheet.getLastRow()
  

  if (!arr[2][0].constructor === Date) { 
        ui.alert("Oops, that's not a date - has the roster layout changed? Please update getCurrentRoster()")
    return false
  }
    
  var packdate = new Date(arr[2][0])
  for (var r = 3; r < arr.length; r++){
    var rosterDetails = arr[r][2]
    if (rosterDetails) {
      var roster = [packdate, getRID(rosterDetails), getRRest(rosterDetails)]
      rosters.push(roster)
    }
  }
  return 
  //sheet.getRange(sheet.getLastRow()+1, 1, rosters.length, rosters[0].length).setValues(rosters)
}

//function convert() {
//  var ss = SpreadsheetApp.getActiveSpreadsheet()
//  var sheet = ss.getSheetByName("sheet28")
//  var range = SpreadsheetApp.getActiveSpreadsheet().getRangeByName("temp_roster")
//  var arr = range.getValues()
//  var last = sheet.getLastRow()
//  var rosters = []
//
//
//  for (var c = 0; c < arr[0].length; c++){
//    var packdate = arr[0][c]
//    for (var r = 2; r < arr.length; r++){
//      var rosterDetails = arr[r][c]
//      if (rosterDetails) {
//        var roster = [packdate, getRID(rosterDetails), getRRest(rosterDetails)]
//        rosters.push(roster)
//      }
//    } 
//  }
//  
//  sheet.getRange(sheet.getLastRow()+1, 1, rosters.length, rosters[0].length).setValues(rosters)
//}



// bug in here - still developing...

function getRID(details){
  var re = /\b(\d{4}\b)/;
  var match = re.exec(details)
  if (match){
    return match[1]
  } else {
    return ""
  }
}

function getRRest(details){
  var re = /(\d{4})\D ?-? ?(.*)/;
  var match = re.exec(details)
  if (match){
    return match[2]
  } else {
    return ""
  }
}

function go(){
  var str = "Julie (nic) 7177"
  Logger.log(getRID(str))
  Logger.log(getRRest(str))
}