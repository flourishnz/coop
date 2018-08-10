//function showAlert() {
//  var ui = SpreadsheetApp.getUi(); // Same variations.
//
//  var result = ui.alert(
//     'Please confirm',
//     'Are you sure you want to continue?',
//      ui.ButtonSet.YES_NO);
//
//  // Process the user's response.
//  if (result == ui.Button.YES) {
//    // User clicked "Yes".
//    ui.alert('Confirmation received.');
//  } else {
//    // User clicked "No" or X in the title bar.
//    ui.alert('Permission denied.');
//  }
//}
//
//function showDialog(data) {
//  var template = HtmlService.createTemplateFromFile('testDialog')
//      .data = data
//  template.setSandboxMode(HtmlService.SandboxMode.IFRAME)
//      .setWidth(400)
//      .setHeight(300)
//  var ui = template.setTitle("Contacts").evaluate()
//  
//  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
//      .showModalDialog(ui, 'Hmmm');
//}
//
//// Create a trigger for the script.
//ScriptApp.newTrigger('myFunction').forSpreadsheet('id of my spreadsheet').onEdit().create();
//Logger.log(ScriptApp.getProjectTriggers()[0].getHandlerFunction()); // logs "myFunction"

//-------------------------
//function getNames(){
//  var ranges = SpreadsheetApp.getActive().getNamedRanges();
//  var names = []
//  for (var i = 0; i < ranges.length; i++) {
//    names.push(ranges[i].getName());
//  }
//  return names
//}
//function testt() {
//  say(getTweaks())
//}


//function getTweaks(){// hack - not function oriented
//  var orders = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Orders")
//                      .getDataRange().getValues()
//  var tweaks = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Pre-tweak Orders")
//                       .getDataRange().getValues()
//  var uptweaks = []
//  var vendor
// 
//  for (var i = 0; i < orders.length; i++){
//    vendor = orders[i][VENDOR_COLUMN-1].trim().toLowerCase()
//    if (vendor === "freshdirect" || vendor === "purefresh" || vendor === "chantal" ) {
//      var order = orders[i]
//      var tweak = tweaks[i]
//      if (order[PRODUCT_COLUMN-1] !== tweak[PRODUCT_COLUMN-1]) {
//        // hunt for matching product because sheets are not aligned...
//        // in the meantime...
//        Logger.log("Quitting: haven't written code to handle misaligned tweak sheet")
//        return
//      }
//      if (order[6] > tweak[6]) {
//        uptweaks.push([order[PRODUCT_COLUMN], order[6]-tweak[6]])
//      }
//      
//    }
//  }
//  return uptweaks
//}

