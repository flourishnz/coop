// v 0.1
// temporary run-once code written to add group b members to Merged sheet (which already had group A members)

function addGroupB() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var newbies = getEmailsGroupB("1L9l65gGWyAatcwuaUFGm06iAgXgkGqMnqzRkdIhd0bk")
  
  
  try {
    ss.addEditors(newbies)
  }
  catch (err) {
    SpreadsheetApp.getUi().alert("Invalid email address: " + member.email.trim())
    Logger.log("Invalid email address: " + member.email.trim() + "\n Worksheet not shared.")
  } 
}


function getEmailsGroupB(id){
  var ssID = id || "1L9l65gGWyAatcwuaUFGm06iAgXgkGqMnqzRkdIhd0bk"
  var ssB = SpreadsheetApp.openById(ssID);
  return ssB.getEditors()
}
