// FRESH - Bin list 
// June 2018 - Use document templates because Apps Script can't create a document with columns yet

/*
createReportBinList
fillBinList(table)
getFreshMembersWhoOrdered
*/

// v1.2 Break every 5 rows instead of 4
// v1.1 Add in Try-Catch error handling around conversion to pdf - may help with server unavailable - maybe not
//      Also returns URL of the pdf if no error, or else the (usually temporary) report document

function createReportBinList(){
  var TEMPLATE_ID = '1suKh6TfvHPsD57FNcvq6gkVXS1Sq7-zzQF8yZm1mvaU'  // bin list template
  var FOLDER_ID = '1KPMf3cnXFJJ7L50zOucPYqaWNoHx2wIp'               // reports go to fresh/reports
  
  var packDate = getPackDateFromFilename()
  var PDF_FILE_NAME = Utilities.formatDate(packDate, "GMT+12:00", "yyyy-MM-dd") + ' Fresh Bin list'
  

  // Set up the docs
  
  var copyFile = DriveApp.getFileById(TEMPLATE_ID).makeCopy(DriveApp.getFolderById(FOLDER_ID)) 
  var doc = DocumentApp.openById(copyFile.getId())
  var body = doc.getBody()
  var header = doc.getHeader()
    
  // Document Heading - set the packdate
  header.replaceText('%PACKDATE%', Utilities.formatDate(packDate, "GMT+12:00", "EEEE, d MMMM yyyy"))

  var newt = body.getTables()[0]
  fillBinList(newt)

  var p = body.appendParagraph(newt.getNumRows() + " bins required")
  p.setAttributes({"FONT_SIZE": 14, "FOREGROUND_COLOR": "#0000ff"})
  say(p.getAttributes())
  
  //------------------------------------------
  // Create PDF from doc, rename it if required and delete the doc
    
  doc.saveAndClose()
  try {
    var pdf = DriveApp.getFolderById(FOLDER_ID).createFile(copyFile.getAs('application/pdf'))
    pdf.setName(PDF_FILE_NAME)
    copyFile.setTrashed(true)
    return pdf.getUrl()
  } 

  catch (err) {
    SpreadsheetApp.getUi().alert("Error while converting Bin list to pdf\n" + err)
    return copyFile.getUrl()
    }   
}



function fillBinList(table) {
  var templateRow = table.getRow(0).removeFromParent()
  //initialise
  var prevId = 1000
  var member = {}
  
   // Get orders
  var data = getFreshMembersWhoOrdered().sort()  // returns [[]]
 
  // create a row for each member
  for (var i=0; i < data.length; i++) {
    member = data[i]
    var newRow = templateRow.copy()
    newRow.replaceText("%id%", member.id)
    newRow.replaceText("%name%", member.name.split()[0])
    
    newRow = table.appendTableRow(newRow)
    
    // put in a break every 5 rows to make it easier to manage the labels
    if ((i+1)%5 == 0) {
      newRow.setMinimumHeight(40)
    }
    prevId = member.id
  }
  
}

function getFreshMembersWhoOrdered(){// returns array of objects
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var headers = isDRY && ss.getRangeByName('ord_Headers').getValues() || ss.getRangeByName('tot_Headers').getValues()
  var members = []
  
  for (var i = 0; i< headers[0].length; i++) {
    if (headers[2][i] === "ordered") {
      members.push({name: headers[0][i], id: headers[1][i]})
    }
  }
  
  return members
}
