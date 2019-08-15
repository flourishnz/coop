// FRESH - Pack Lists 
// June 2018 - Use document templates because Apps Script can't create a document with columns yet

// v1.2 Add in Try-Catch error handling around conversion to pdf - may help with server unavailable - maybe not
//      Also returns URL of the pdf if no error, or else the (usually temporary) report document

// v 1.1 Moved sharePdfPacksheets to Code

function createReportFreshPacklist() {
  
  var TEMPLATE_ID = '1rMSRbBIfGI08ww6pG2E5EKR9NiFrqHRUs198AQrWsE0'  // fresh pack list template
  var FOLDER_ID = '1KPMf3cnXFJJ7L50zOucPYqaWNoHx2wIp'               // reports go to fresh/reports
  
  var packDate = getPackDateFromFilename()
  var PDF_FILE_NAME = Utilities.formatDate(packDate, "GMT+12:00", "yyyy-MM-dd") + ' Fresh packlist'
  
  if (TEMPLATE_ID === '') {   
    SpreadsheetApp.getUi().alert('TEMPLATE_ID needs to be defined in code.gs')
    return
  }

  // Set up the docs and the spreadsheet access
  
  var copyFile = DriveApp.getFileById(TEMPLATE_ID).makeCopy(DriveApp.getFolderById(FOLDER_ID)) 
  var doc = DocumentApp.openById(copyFile.getId())
  var body = doc.getBody()
  var header = doc.getHeader()
  
  // Get orders
  var prodOrders = getOrders()  

  // get a copy of templated tables and remove them from the front of the document in the same command
  var templateTable = body.getTables()[0].removeFromParent()
  var templateFoot = body.getTables()[0].removeFromParent()
  
 // Document Heading - set the packdate
  header.replaceText('%PACKDATE%', Utilities.formatDate(packDate, "GMT+12:00", "EEEE, d MMMM yyyy"))

  var newt = body.appendTable()
//  makeBinList(getFreshMembersWhoOrdered(), newt)
//   doc.saveAndClose()
//   return
   //----------------------------
  var prevVendor = ""
  
  // table for each product
  for (var i = 0; i < prodOrders.length; i++){
    var p = prodOrders[i]
    //if (p.vendor == "Chantal") {continue} // skip some vendors for testing
    var thisTable = templateTable.copy()
    
    // adjust table headers
    thisTable.replaceText("%VENDOR%", p.vendor)
    thisTable.replaceText("%PRODUCT%", p.product)


    var crates = p.total/p.crateCount  
    var numCrates = (crates > 1.4 && crates < 2) || (crates > 2.3 && crates < 3)   //... move this calculation into the product or look up actual order
                       ? Math.floor(crates) +1   //equiv to roundup
                       : Math.round(crates)
 
    
    if (p.crateCount == 1) {
      var crateTxt = numCrates + " " + ((p.unit == "KG") ? "kg expected" : "items expected")
    } else {
      crateTxt = numCrates + ((numCrates == 1) ? " crate" : " crates") + " of " + p.crateCount  + ((p.unit == "KG") ? " kg" : " items") + " expected"
    }
    
    thisTable.replaceText("%CRATES%", crateTxt)
    
    thisTable.replaceText("%TOTAL%", p.total)
    thisTable.replaceText("%UNIT%", ((p.unit == "KG") ? "kg allocated" : "allocated"))
    fillTable(thisTable, p)
    
    // add footer and throw page before appending table to new page
    // unless first product or vendor is Purebread
    
    var groupedVendor = (p.vendor == prevVendor && 
                   (p.vendor == "The Egg Shed" || p.vendor == "Purebread" || p.vendor == "Common Property" ))
    if (i > 0 && !groupedVendor) {
      body.appendTable(templateFoot.copy())
      body.appendPageBreak()
    }
    body.appendTable(thisTable)
    prevVendor = p.vendor
  }
  body.appendTable(templateFoot.copy()) // add footer to last product in the report
  //------------------------------------------
  // Create PDF from doc, rename it if required and delete the doc
    
  doc.saveAndClose()

  try {
    var pdf = DriveApp.getFolderById(FOLDER_ID).createFile(copyFile.getAs('application/pdf')) 
    if (PDF_FILE_NAME !== '') {
      pdf.setName(PDF_FILE_NAME)
    }
    //sharePdfPacksheets(pdf)
    copyFile.setTrashed(true)
    return pdf.getUrl()
  }
  catch(err) {
    SpreadsheetApp.getUi().alert("Error while converting pack list to pdf\n" + err)
    return copyFile.getUrl()
    }
  
}

function fillTable(table, product) {
  var prevQty = 0 
  
  // copy and remove the template row
  var lastRow = table.getRow(table.getNumRows()-1).removeFromParent()  

  // sort all? weighed? produce
//  if (product.unit == "KG") {
    product.orders = _.sortBy(product.orders, ['qty'])
//  }  
  
  // create a row for each order
  for (var i = 0; i < product.orders.length; i++) {
    var order = product.orders[i]
    var newRow = lastRow.copy()
    newRow.replaceText("%tweak%", (order.tweak == 0 ? "" : order.tweak))
    newRow.replaceText("%name%", order.name.split(" ")[0])
    newRow.replaceText("%bin%", order.id)
    newRow.replaceText("%qty%", order.qty)

    newRow = table.appendTableRow(newRow)
    
 
    // leave a gap between quantities
    if (order.qty != prevQty) {
      prevQty = order.qty
      if (i > 0) {newRow.setMinimumHeight(40)}
    } 
  }
  // table.appendTableRow().setMinimumHeight(100).appendTableCell().appendHorizontalRule()  
}


//  VENDOR_COLUMN = 2
//  PRODUCT_COLUMN = 3
//  UNIT_COLUMN = 4
//  CRATE_COUNT_COLUMN = 5
//  PRICE_EXCL_COLUMN = 6
//  PRICE_COLUMN = 7
//  TOTAL_KGS_COLUMN = 8
//  TOTAL_CRATES_COLUMN = 9
//
//  FIRST_ORDER_COLUMN = 10
//  FIRST_ORDER_ROW = 5
//  
//  USERNAME_ROW = 3
//  USERID_ROW = 4



function getOrders(){
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var data = ss.getSheetByName('Orders').getDataRange().getValues()
  var products = []
  var preTweaked = ss.getSheetByName("Pre-tweak Orders").getDataRange().getValues()
  
  //  { product: "Oranges", price:3.95, orders: {[name: "Fred", qty: 2, id: '9059'],
  //                                            [name: "Wilma", qty: 5, id: '9023']}}
  
  // for each product...
  for (var product = FIRST_ORDER_ROW-1; product < data.length; product++) {
    if (data[product][PRODUCT_COLUMN] !== ""){
      var orders = []
      // collect each member's order...
      for (var member = FIRST_ORDER_COLUMN-1; member < data[product].length; member++) {
        if (data[product][member]>0) {
          orders.push({'tweak': rounded(data[product][member] - preTweaked[product][member]),
                       'id': data[USERID_ROW-1][member],
                       'name': data[USERNAME_ROW-1][member],
                       'qty': rounded(data[product][member])}
                     )
        }
      }
      // ...and attach the orders to the product details (still need to add in other available fields)
      if (orders.length > 0) {
        products.push({'product': data[product][PRODUCT_COLUMN-1],   
                       'vendor': data[product][VENDOR_COLUMN-1],
                       'unit': data[product][UNIT_COLUMN-1].toUpperCase(),
                       'crateCount': data[product][CRATE_COUNT_COLUMN-1],
                       'total': rounded(data[product][TOTAL_KGS_COLUMN-1]),
                       'orders': orders
                      })
      }
    }
  }
  return products
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
