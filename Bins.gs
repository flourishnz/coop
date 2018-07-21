function testBL(){
  var TEMPLATE_ID = '1suKh6TfvHPsD57FNcvq6gkVXS1Sq7-zzQF8yZm1mvaU'  // bin list template
  var FOLDER_ID = '1KPMf3cnXFJJ7L50zOucPYqaWNoHx2wIp'               // reports go to fresh/reports
  
  var packDate = getPackDateFromFilename()
  var PDF_FILE_NAME = Utilities.formatDate(packDate, "GMT+12:00", "yyyy-MM-dd") + ' Fresh Bin list'
  

  // Set up the docs
  
  var copyFile = DriveApp.getFileById(TEMPLATE_ID).makeCopy(DriveApp.getFolderById(FOLDER_ID)) 
  var doc = DocumentApp.openById(copyFile.getId())
  var body = doc.getBody()
  var header = doc.getHeader()
  
// debugging / exploring
  var x = body.getChild(1)
  say (x)
  say(body.getChildIndex(x))
 // say(body.getChildIndex(body.getTables[1]))
  
  
  
  // get a copy of templated tables and remove them from the front of the document in the same command
  //templateTable = body.getTables()[0].removeFromParent()
  //var templateFoot = body.getTables()[0].removeFromParent()
  
  // Document Heading - set the packdate
  header.replaceText('%PACKDATE%', Utilities.formatDate(packDate, "GMT+12:00", "EEEE, d MMMM yyyy"))

  var newt = body.getTables()[0]
  fillBinList(newt)

  var p = body.appendParagraph(newt.getNumRows() + " bins required")
  p.setAttributes({"FONT_SIZE": 14, "FOREGROUND_COLOR": "#0000ff"})
  say(p.getAttributes())
  
}



function fillBinList(table) {
  var templateRow = table.getRow(0).removeFromParent()
  var prevId = 1000
  var member = {}
  
   // Get orders
  var data = getMembersWhoOrdered()  
 
  // create a row for each member
  for (var i=0; i < data.length; i++) {
    member = data[i]
    var newRow = templateRow.copy()
    newRow.replaceText("%id%", member.id)
    newRow.replaceText("%name%", member.name.split()[0])
    
    newRow = table.appendTableRow(newRow)
    
    // put in a break every 4 rows to make it easier to manage the labels
    if ((i+1)%4 == 0) {
      newRow.setMinimumHeight(40)
    }
    prevId = member.id
  }
  
}


