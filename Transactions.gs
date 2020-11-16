/* TRANSACTIONS

v 0.1

Getting banking transactions
============================

getTransactions()           - returns all/some banking transactions, may call defineTransactions
getTransactions(id) 
getTransactions(id, limit)  eg getTransactions(7177, 2) may return [{date: nnn, amount: 300},{date: nnn, amount: 300}]

defineTransactions() stores all baking transaction in global varaiable TRANSACTIONS for quick access by getTransactions
                           
*/

function quicktest(){
  var x = getTransactions(7177, 3)
  say(x)
}

function getTransactions(id="", limit=100) {
  if (typeof TRANSACTIONS == 'undefined'){
    defineTransactions()
  }
  
  if (id == ""){return TRANSACTIONS}
  if (typeof TRANSACTIONS[id] == 'undefined') {return[]}
  return TRANSACTIONS[id].slice(0, limit)

}


function defineTransactions(){
  // defines/updates global variable TRANSACTIONS
  // latest transaction for id 7000 is  TRANSACTIONS[7000][0] => {date: dsads, amount: 3232.34}
  
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var data = ss.getSheetByName("Banking").getDataRange().getValues()
  var idCol = ss.getRangeByName("bank_Bin").getColumn()-1
  var amountCol = ss.getRangeByName("bank_Amount").getColumn()-1
  TRANSACTIONS = []
  
  for (var i = 0; i < data.length; i++) {
    var id = data[i][idCol]
    if (!isValidId(id)) {continue}
    
    if (!(id in TRANSACTIONS)) {TRANSACTIONS[id] = []}
    TRANSACTIONS[id].unshift({date: data[i][0], 
                              amount: data[i][amountCol]
                             })
  }
}
