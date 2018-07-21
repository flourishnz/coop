function convert() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var sheet = ss.getSheetByName("sheet28")
  var range = SpreadsheetApp.getActiveSpreadsheet().getRangeByName("temp_roster")
  var arr = range.getValues()
  var last = sheet.getLastRow()
  var rosters = []


  for (var c = 0; c < arr[0].length; c++){
    var packdate = arr[0][c]
    for (var r = 2; r < arr.length; r++){
      var rosterDetails = arr[r][c]
      if (rosterDetails) {
        var roster = [packdate, getRID(rosterDetails), getRRest(rosterDetails)]
        rosters.push(roster)
      }
    } 
  }
  
  sheet.getRange(sheet.getLastRow()+1, 1, rosters.length, rosters[0].length).setValues(rosters)
}




function getRID(details){
  var re = /^(\d{4})\D ?-? ?(.*)/;
  var match = re.exec(details)
  if (match){
    return match[1]
  } else {
    return ""
  }
}

function getRRest(details){
  var re = /^(\d{4})\D ?-? ?(.*)/;
  var match = re.exec(details)
  if (match){
    return match[2]
  } else {
    return ""
  }
}

function go(){
  Logger.log(getRRest("1234- abcd f"))
}