// Membership Report
// v0.2  Move email addresses to globals
// v0.1 under development

function createReportMembers() {
  // get data
  var members = getMembers()
  var data = []
  
  for (var i = 0; i < members.length; i++) {              //20; i++){ //      members.length; i++) {
    var m = members[i]
    var lastPayment = m.getLatestPayment()
    data.push([m.id, m.name, rounded(m.getCurrentBalance()), lastPayment.date, lastPayment.amount])
  }
  
  data.sort(function (a, b) {return a[2] < b[2] ? -1 : a[2] > b[2] ? 1 : 0})
  
//  // open file
//  var packDate = getPackDateFromFilename()
//  var filename = Utilities.formatDate(packDate, "GMT+12:00", "yyyy-MM-dd") + ' Membership Report'
//  var doc = DocumentApp.create(filename)
//  var body = doc.getBody()
//  
//  // write doc
//  body.appendTable(data)
//  
//  // close and notify
//  var id = doc.getId()
//  tellIT('Created doc with id: ' + id)
//  doc.saveAndClose()
  return
}