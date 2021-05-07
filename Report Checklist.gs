// Developing DRY - Feb Jun 2018

// v0.74 Rename reports
// v0.4


function createReportChecklist() {
  
  var TEMPLATE_ID = '1KXMSjY6iFZjHoC20n9HJ26TeFY0T7CiIxp0A9OvDL-s'   // pickup checklist template
  var FOLDER_ID = '1Ur9LaAUeYzlIFxQ77bO3oj0hJ0UIDULc'                // reports go to dry/reports
  var packDate = getPackDateFromFilename()
  var PDF_FILE_NAME = Utilities.formatDate(packDate, "GMT+12:00", "yyyy-MM-dd") + ' Checklist'

  // Set up the docs and the spreadsheet access
  
  var copyFile = DriveApp.getFileById(TEMPLATE_ID).makeCopy(DriveApp.getFolderById(FOLDER_ID)) 
  var doc = DocumentApp.openById(copyFile.getId())
  var body = doc.getBody()
  var header = doc.getHeader()
  
  // Get data

  var data = getMembersWhoOrdered()
 
  // Document Heading - set the packdate
 header.replaceText('%DATE%', Utilities.formatDate(packDate, "GMT+12:00", "EEEE, d MMMM yyyy"))
  
  // Create table
 
  var table = body.appendTable(data)
  table.setColumnWidth(0, 25)           // for tick box
  table.setColumnWidth(2, 65)           // for ID number
  table.setBorderWidth(0)               // no table borders
  


  // Format headers
  
  var headerStyle = {};
  headerStyle[DocumentApp.Attribute.BACKGROUND_COLOR] = '#444444';  
  headerStyle[DocumentApp.Attribute.BOLD] = true;  
  headerStyle[DocumentApp.Attribute.FOREGROUND_COLOR] = '#FFFFFF';
  headerStyle[DocumentApp.Attribute.VERTICAL_ALIGNMENT] = DocumentApp.VerticalAlignment.CENTER;
  

  var headers = table.getRow(0)  
  for (var j=0; j < headers.getNumCells(); j++) {
    headers.getCell(j).setAttributes(headerStyle)
  }
 
  
  // format data
  var bodyStyle = {};
  bodyStyle[DocumentApp.Attribute.VERTICAL_ALIGNMENT] = DocumentApp.VerticalAlignment.CENTER;

  for (var i = 1; i < table.getNumRows(); i++){
    var row = table.getRow(i)
    row.getCell(0).setText(String.fromCharCode(9744)).editAsText().setFontSize(12)    // tick box
    for (var j=0; j < row.getNumCells(); j++) {
      row.getCell(j).setAttributes(bodyStyle)
    }
  }
  
  
  //------------------------------------------
  // Create PDF from doc, rename it if required and delete the doc
    
  doc.saveAndClose()
  var pdf = DriveApp.getFolderById(FOLDER_ID).createFile(copyFile.getAs('application/pdf'))  

  if (PDF_FILE_NAME !== '') {
    pdf.setName(PDF_FILE_NAME)
  } 
  
  copyFile.setTrashed(true)
  
  
}


function getMembersWhoOrdered(){// returns array
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var headers = isDRY && ss.getRangeByName('ord_Headers').getValues() || ss.getRangeByName('tot_Headers').getValues()
  var data = []
  
  for (var jj = 0; jj< headers[0].length; jj++) {
    if (headers[2][jj] === "ordered") {
      data.push(['', headers[0][jj], headers[1][jj]])
    }
  }
  data.sort().unshift(["", "Member", "Account Number"])
  return data
}

//========================================
// from code by // dev: andrewroberts.net
// Demo script - http://bit.ly/createPDF
// var TEMPLATE_ID = '1wtGEp27HNEVwImeh2as7bRNw-tO4HkwPGcAsTrSNTPc' // Demo template
/**  
 * Take the fields from the active row in the active sheet
 * and, using a Google Doc template, create a PDF doc with these
 * fields replacing the keys in the template. The keys are identified
 * by having a % either side, e.g. %Name%.
 *
 * @return {Object} the completed PDF file
 */
  //SpreadsheetApp.getUi().alert('New PDF file created in the root of your Google Drive')
