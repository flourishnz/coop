// DRY - Pack Lists 
// June 2018 - Use document templates because Apps Script can't create a document with columns yet

// v0.06

function createReportDryPickupLists() {
  
  var TEMPLATE_ID = '1IQKKvDM8AMvvxi5VsNmARzzFdiWWk3nstC6POTQzDNQ'  // dry pack list template
  var FOLDER_ID = '1Ur9LaAUeYzlIFxQ77bO3oj0hJ0UIDULc'               // reports go to dry/reports
  
  var packDate = getPackDateFromFilename()
  var PDF_FILE_NAME = Utilities.formatDate(packDate, "GMT+12:00", "yyyy-MM-dd") + ' Dry Pick-up Lists'
  
  if (TEMPLATE_ID === '') {   
    SpreadsheetApp.getUi().alert('TEMPLATE_ID needs to be defined in code.gs')
    return
  }
  
  // Set up the docs and the spreadsheet access
  
  var copyFile = DriveApp.getFileById(TEMPLATE_ID).makeCopy(DriveApp.getFolderById(FOLDER_ID)) 
  var doc = DocumentApp.openById(copyFile.getId())
  var body = doc.getBody()
  var header = doc.getHeader()
    
  // get a copy of templated tables and clear the document
  var templateMemberTable = body.getTables()[0]
  var templateOrdersTable = body.getTables()[1]
  body.clear()
  
  // copy and remove the template order row
  var dataRow = templateOrdersTable.getRow(1).removeFromParent()
  
  // Document Heading - set the packdate
  header.replaceText('%PACKDATE%', Utilities.formatDate(packDate, "GMT+12:00", "EEEE, d MMMM yyyy"))

  // Get orders
  var memberOrders = getDryOrdersByMember().reverse()

  
  //------------------------------------------------------------------------------------------------
  // for each member, on a new page, add a member header and then a table for each type of product
  //------------------------------------------------------------------------------------------------
  
  for (var i = 0; i < memberOrders.length; i++){
    var member = memberOrders[i]
    
    if (i>0) {body.appendPageBreak()}

    // create and insert member header
    var memberTable = templateMemberTable.copy()
    memberTable.replaceText("%member%", member.name)
    memberTable.replaceText("%id%", member.id)
    memberTable = body.appendTable(memberTable)
    
    
    // create a separate table for each measuring-type of product
    var weighables = templateOrdersTable.copy().replaceText("%type%", "Weighables (kg)")
    var countables = templateOrdersTable.copy().replaceText("%type%", "Countables")
    var pourables  = templateOrdersTable.copy().replaceText("%type%", "Pourables (litres)")
        
    // add a row for each order to one of the tables
    for (var j = 0; j < member.orders.length; j++) {
      var order = member.orders[j]
      var newRow = dataRow.copy()
      
      // newRow.replaceText("%tweak%", (order.tweak == 0 ? "" : order.tweak))
      newRow.replaceText("%product%", order.product)
      newRow.replaceText("%qty%", order.qty)
      
      if (order.unit == "kg" || order.product.toLowerCase() == "coconut oil") {
        newRow = weighables.appendTableRow(newRow)
      } else if (order.unit == "litre") {
        newRow = pourables.appendTableRow(newRow)
      } else {
        newRow = countables.appendTableRow(newRow)                  
      }
    }    
    if (countables.getNumRows() > 1) {body.appendTable(countables)}
    if (weighables.getNumRows() > 1) {body.appendTable(weighables)}
    if (pourables.getNumRows() > 1) {body.appendTable(pourables)}
  }

  //------------------------------------------
  // Create PDF from doc, rename it if required and delete the doc
    
  doc.saveAndClose()
  
  var pdf = DriveApp.getFolderById(FOLDER_ID).createFile(copyFile.getAs('application/pdf'))  
  if (PDF_FILE_NAME !== '') {
    pdf.setName(PDF_FILE_NAME)
  } 
  sharePdfPacksheets(pdf)
  
  copyFile.setTrashed(true)

}




function getDryOrdersByMember(){// actually works with Fresh except haven't handled TWEAKS (OR LACK THEREOF)
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var data = ss.getSheetByName('Orders').getDataRange().getValues()
  var members = []
  //if (isFresh()) {var preTweaked = ss.getSheetByName("Pre-tweak Orders").getDataRange().getValues()}
  
  //  { name: "Fred", id: "7098", orders: {{product: "Oranges",    unit: "kg", qty:3, price: 4.55},
  //                                       {product: "Watermelon", unit: "ea", qty:2, price: 8.12}}
  
  // for each member...
  for (var member = FIRST_ORDER_COLUMN-1; member < data[0].length; member++) {
    if (data[USERID_ROW-1][member] !== ""){   // contains member id
      var orders = []
      
      // collect each member's orders...
      for (var product = FIRST_ORDER_ROW-1; product < data.length; product++) {
        if (data[product][member]>0) {
          orders.push({     //'tweak': rounded(data[product][product] - preTweaked[product][member]),
                       'product': data[product][PRODUCT_COLUMN-1],
                       'unit': data[product][UNIT_COLUMN-1].toLowerCase(),
                       'price': data[product][PRICE_COLUMN-1],
                       'qty': rounded(data[product][member])}
                     )
        }
      }
      
      // ...and attach the orders to the member details 
      if (orders.length > 0) {
        members.push({'name': data[USERNAME_ROW-1][member],   
                       'id': data[USERID_ROW-1][member],
                       'orders': orders
                      })
      }
    }
  }
  
  return _.sortBy( members, ['name'])
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
