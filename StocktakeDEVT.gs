//// STOCKTAKE
//
//// v1.0 Breaking report into groups - UNDER DEVELOPMENT
//// v0.1
//
//
//function createReportStocktake() {
//  
//  var TEMPLATE_ID = '1jww1efQoKL1KWcRuNSZcdmAZsyfyy2VO8P7M-c7KIXU'  // dry stocktake template
//  var FOLDER_ID = '1Ur9LaAUeYzlIFxQ77bO3oj0hJ0UIDULc'               // reports go to dry/reports
//  
//  var packDate = getPackDateFromFilename()
//  var PDF_FILE_NAME = Utilities.formatDate(packDate, "GMT+12:00", "yyyy-MM-dd") + ' Stocktake List'
//  
//  // Set up the docs and the spreadsheet access
//  
//  var copyFile = DriveApp.getFileById(TEMPLATE_ID).makeCopy(DriveApp.getFolderById(FOLDER_ID)) 
//  var doc = DocumentApp.openById(copyFile.getId())
//  var body = doc.getBody()
//  var header = doc.getHeader()
//  
//  // get a copy of templated tables and clear the document
//  var templateOrdersTable = body.getTables()[0]
//  body.clear()
//  
//  // copy and remove the template order row
//  var dataRow = templateOrdersTable.getRow(1).removeFromParent()
//  
//  // Document Heading - set the packdate
//  header.replaceText('%PACKDATE%', Utilities.formatDate(packDate, "GMT+12:00", "EEEE, d MMMM yyyy"))
//
//  // create a separate table for each measuring-type of product
//  var weighables = templateOrdersTable.copy().replaceText("%type%", "Weighables (kg)")
//  var countables = templateOrdersTable.copy().replaceText("%type%", "Countables")
//  var pourables  = templateOrdersTable.copy().replaceText("%type%", "Pourables (litres)")
//
//  // Get products
//  var products = getProducts()
//
//  // init
//  var currGroup = ' '
//  
//  // add a row for each product to one of the tables
//  for (var i = 0; i < products.length; i++){
//    var product = products[i]
//    var newRow = dataRow.copy()
//      
//    // newRow.replaceText("%tweak%", (order.tweak == 0 ? "" : order.tweak))
//    newRow.replaceText("%product%", product.product)
//    
//    if (product.unit == "kg" || product.product.toLowerCase() == "coconut oil") {
//      if (product.group !== currGroup){
//        currGroup = product.group
//        weighables.appendTableRow([product.group])
//      }
//      newRow = weighables.appendTableRow(newRow)
//    } else if (product.unit == "litre") {
//      newRow = pourables.appendTableRow(newRow)
//    } else {
//      newRow = countables.appendTableRow(newRow)                  
//    }
//  }
//  
//  if (countables.getNumRows() > 1) {body.appendTable(countables)}
//  body.appendPageBreak()
//  if (weighables.getNumRows() > 1) {body.appendTable(weighables)}
//  if (pourables.getNumRows() > 1) {body.appendTable(pourables)}
//
//
//  //------------------------------------------
//  // Create PDF from doc, rename it if required and delete the doc
//    
//  doc.saveAndClose()
//  
//  var pdf = DriveApp.getFolderById(FOLDER_ID).createFile(copyFile.getAs('application/pdf'))  
//  if (PDF_FILE_NAME !== '') {
//    pdf.setName(PDF_FILE_NAME)
//  } 
//  //sharePdfPacksheets(pdf)   // DON"T SHARE WHILE TESTING
//  copyFile.setTrashed(true)
//}
//
//
//function getProducts(){
//  var ss = SpreadsheetApp.getActiveSpreadsheet()
//  var data = ss.getSheetByName('Orders').getDataRange().getValues()
//  var products = []
//  var group = ''
//      
//  // collect each product
//  for (var row = FIRST_ORDER_ROW-1; row < data.length; row++) {
//    if (data[row][GROUP_COLUMN] !== "") {group = data[row][GROUP_COLUMN]}
//    if (data[row][PRODUCT_COLUMN] !== "") {
//      products.push({'product': data[row][PRODUCT_COLUMN-1],
//                   'unit': data[row][UNIT_COLUMN-1].toLowerCase(),
//                   'price': data[row][PRICE_COLUMN-1],
//                   'group': group
//                    })}}
//  return products
//}
