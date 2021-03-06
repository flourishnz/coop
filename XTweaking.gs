// TWEAKING
// v 1.5 Tweaking target changed because Chantal now have minimums
// v 1.41  Call tellIT instead of tellJulie
// v 1.4 playing with success message after generating reports in runFreshReports
// v 1.3 doneTweaking: generate reports
// v 1.2 Modify summariseThis to rounding to nearest kg when order is below 0.8 crates, instead of nearest crate
//       and ease test for whether tweaking required to be within 50g of target.

/*
startTweaking
doneTweaking
runFreshReports
runDryReports
createSheetPreTweaks
getProduct
saveProductOrders
summariseOrders
showTweakbar
summariseThis
tweakAdd
tweakScale
add
*/


function startTweaking(){
  closeOrdering("Closed - Tweaking")  //lock and setStatus 
  createSheetPreTweaks()
  // summarise orders - group by tweak up, tweak down, tweak out
  // email summary?
  // start with (or ignore) products that have multiple varieties
  // confirm tweak outs
  // consider adding varieties
  
  // Tweak a product
  //   Summarise total, excess or shortfall
  //   Calculate scaled increase and equal increase
  
}


function doneTweaking(){
  setStatus("Closed - Tweaked")
  runFreshReports()
  tellIT("Reports created successfully.")
}

function runFreshReports(){
  var pckl = createReportFreshPacklist()
  var binl = createReportBinList() 
  if (pckl || binl) {
    varcreateReportChecklist
    msg += pckl ? "Created " + pckl + "\n" 
                : "Pack list not created\n"
    msg += binl ? "Created " + binl + "\n" 
                : "Bin list not created\n"    
    SpreadsheetApp.getUi().alert(msg)
  }

}

function runDryReports(){
  createReportDryPickupLists()
  createReportPickupChecklist()
  createReportStocktake()
}


function createSheetPreTweaks(){
  var name = "Pre-tweak Orders"
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var orders = ss.getSheetByName('Orders')
  var initialActiveSheet = ss.getActiveSheet()    // so that we can return to this sheet after creating the pre-tweaks sheet
  try {
    if (ss.getSheetByName(name) !== null ){ss.deleteSheet(ss.getSheetByName(name)) }
    var sheet = ss.insertSheet(name, orders.getIndex() ,{template: orders})    //getIndex is 1-based, insert-sheeet is 0-based! so this inserts sheet immediately after Orders  
    initialActiveSheet.activate()                // because inserting a sheet changes the active sheet to the new sheet
    lockSheet(sheet, "Pre-tweaks")
  }

  catch (err) {
    Logger.log(err)
  }
}


function getProduct(arg){// argument is either row name or name of product
  var ss = SpreadsheetApp.getActive()
  var data = ss.getRangeByName("ord_FreshDirect_Orders").getValues()
  
  var row = isNumeric(arg) && arg
  var productName = row && ss.getSheetByName("Orders").getRange(row, PRODUCT_COLUMN).getValue() || arg
  
  var prodData = ArrayLib.filterByText(data, PRODUCT_COLUMN - 1, productName)
  var obj = {product: prodData[0][PRODUCT_COLUMN-1] ,
             vendor: prodData[0][VENDOR_COLUMN-1],
             unit: prodData[0][UNIT_COLUMN-1],
             crateCount: prodData[0][CRATE_COUNT_COLUMN-1],
             priceExcl: prodData[0][PRICE_EXCL_COLUMN-1],
             price: prodData[0][PRICE_COLUMN-1],
             totalKgs: prodData[0][TOTAL_KGS_COLUMN-1],
             totalCrates: prodData[0][TOTAL_CRATES_COLUMN-1],
             orders: prodData[0].splice(FIRST_ORDER_COLUMN-1, prodData[0].length-FIRST_ORDER_COLUMN)
            }
  obj.summary = summariseOrders(obj.orders)
  Logger.log(obj)
  return obj
}

function saveProductOrders(product){// save to CURRENT location, if matches. Maybe should just seek it
  var sheet = SpreadsheetApp.getActiveSheet()
  var thisRow = sheet.getActiveCell().getRow()
  if (sheet.getRange(thisRow, PRODUCT_COLUMN).getValue() == product.product) {
    var range = sheet.getRange(thisRow, FIRST_ORDER_COLUMN, 1, product.orders.length)
  
    //set Range
    range.setValues([product.orders])
  }
}





//function getProducts(){
//  var data = getPureFresh()
//   var prodlist = []
//  for (var i=0; i<data.length; i++){
//    if (data[i][8]%1 > 0){   //modulo 0
//      obj = {product: data[i][2],
//          
//             orders: data[i].splice(FIRST_ORDER_COLUMN-1, data[i].length-FIRST_ORDER_COLUMN-1).sort()}
//      //Logger.log(JSON.stringify(obj, null, 4))
//      prodlist.push(obj)
//    }
//  }
//  return prodlist
//}


//-----------another approach

function summariseOrders(orders){// selects and sorts non-zero entries 
  var summary = []
  for (var i=0; i<orders.length; i++){
    if (orders[i]){
      summary.push(orders[i])
    }
  }
  summary.sort()
  return summary
}


function showTweakbar(data){
  var ui = SpreadsheetApp.getUi()
  var template = HtmlService.createTemplateFromFile('tweakbar');
  template.data = data
  var sb = template.evaluate()
                   .setTitle('Order details');
  ui.showSidebar(sb);
}


function summariseThis() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var thisrow = sheet.getActiveCell().getRow();
  var product = getProduct(thisrow)
  

  if (product.totalKgs < 0.55*product.crateCount) {
    //  roundKg
    var total = Math.round(product.summary.reduce(add, 0)*10000)/10000
    var target = 0  //Math.round(product.totalKgs)
    } 
  else {
    //  round crate
    var total = Math.round(product.summary.reduce(add, 0)*10000)/10000
    var target = product.crateCount * Math.round(product.totalCrates)
  }
  var shortfall = target - total
  var count = product.summary.length
  var unit = (product.unit == "kg" ? " kg" : " items")
  
  //summary headers
  var data = []
  data.push(product.product)
  data.push(count + " orders totalling " + total + unit)
  data.push("Target: " + target) 
  if (total == target || (total >= target - 0.05 && total <= target + 0.05)) {
    data.push("NO TWEAKING REQUIRED.")
  } else {
    data.push((shortfall > 0 ? "Short by " + rounded(shortfall) : "Over by " + rounded(shortfall*-1)) + unit)
    data.push("Change per order: " + (product.unit == "kg" ? 
              Math.round(shortfall/count *1000) + " g" : 
              rounded(shortfall/count)
    ))
    data.push("Percentage change reqd: " + Math.round(shortfall/total *100) + "%")
  }
  data.push(" ")

  // summary orders
  var unique = ArrayLib.unique(product.summary)
  for (var i=0; i< unique.length; i++){
    data.push(ArrayLib.countif(product.summary, unique[i], true) + " orders of "  + unique[i] + " kg")
 }
  
 say(data)
  showTweakbar(data)
}


// add to all orders of this product, the same amount, in grams, +ve or -ve
function tweakAdd(){
  var sheet = SpreadsheetApp.getActiveSheet();
  var thisrow = sheet.getActiveCell().getRow();
  var product = getProduct(thisrow)
  var ui = SpreadsheetApp.getUi()


  // Get adjustment (positve or negtive ok, get in grams, convert to kg)
  
  var result = ui.prompt(product.product, "Adjust all orders by... (g)",ui.ButtonSet.OK_CANCEL)
  var button = result.getSelectedButton();
  var adjustment = parseFloat(result.getResponseText())/1000
  
  // Add to each non-zero order
  if ((button == ui.Button.OK)  && isNumeric(adjustment) && adjustment != 0)  {
    for (var i=0; i < product.orders.length; i++) {
      if (product.orders[i] > 0 && (product.orders[i] + adjustment > 0)) {
        product.orders[i] = product.orders[i]+adjustment
      }
    }
  }
  saveProductOrders(product)
  summariseThis()
}


// scale all orders of this product by the same amount, in %, +ve or -ve
function tweakScale(){
  var ui = SpreadsheetApp.getUi()
  var sheet = SpreadsheetApp.getActiveSheet();
  var thisrow = sheet.getActiveCell().getRow();
  var product = getProduct(thisrow)
  var qty

  // Get adjustment (positive or negtive ok, get in 2 digit percentage, convery to decimal form,round results)
  
  var result = ui.prompt(product.product, "Scale all orders by... (%)", ui.ButtonSet.OK_CANCEL)
  var button = result.getSelectedButton();
  var adjustment = parseFloat(result.getResponseText())/100   //eg 35% becomes 0.35
  
  // Scale each non-zero order
  if ((button == ui.Button.OK)  && isNumeric(adjustment) && adjustment != 0)  {
    for (var i=0; i < product.orders.length; i++) {
      
      if (product.orders[i] > 0 && (product.orders[i] + adjustment > 0)) {
        qty = product.orders[i]
        qty = Math.round((qty + qty*adjustment)*20)/20
        product.orders[i] = qty
      }
    }
  }
  saveProductOrders(product)
  summariseThis()
}

function add(a, b) { // usage: var sum = [1, 2, 3].reduce(add, 0);
    return a + b;
}
