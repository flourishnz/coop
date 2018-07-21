// STATEMENTS
// v0.2

// Developing May-June 2018

//function showDialog(data) {
//  var template = HtmlService.createTemplateFromFile('testDialog')
//      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
//      .setWidth(400)
//      .setHeight(300);
//      .data = data;
//  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
//      .showModalDialog(html, 'My custom dialog');
//}

//...still to do 
// updateEmail
// updateID ?


function testStatements(){
//  var id = "3102"
//  var member = getLatestTotals(id)
//  say (member)
//  say(getBanking(member.id))
//  say (getEmail(member.id))
  var rtn = getFreshContacts()
  var data = "Got " + rtn.length + " contacts"
  showDialog(data)
}


function testMC(){
  //var members = getMembers()
  var member = {name: "my test", id: "9998", email: "about@example.com", homePhone: 432432,               address: "321 Lets Drive"}
//  var member = {name: "my test", id: "9998", email: "about@example.com", mobile: "123", address: "321 Lets Drive"}
  addMemberToContacts(member)

  var member = {name: "my test", id: "9998", email: "about@example.com", mobile: "123", address: "321 Lets Drive"}
  addMemberToContacts(member)
}


function addMemberToContacts(member) {
  var coopGroup = ContactsApp.getContactGroup("Co-op members")

  if (!coopGroup){
    log(["Cannot access Co-op Contacts from this account", member])
    return
  }
  
  var contacts = ContactsApp.getContactsByName(member.name)
  if (contacts.length == 0) {
    addContact(member, coopGroup)
  } else if (contacts.length == 1){
    updateContact(contacts[0], member, coopGroup)
  } else {//... handle this better - alert - etc
    log(["Multiple contacts exist with this name. Not updated", member])
  }
}


function addContact(member, group){
//    log(["Request to add contact", member, "just logging requests"])
//    return
  var firstName = member.name.split(" ")[0]
  var theRest = member.name.substring(firstName.length+1, member.name.length)
  var contact = ContactsApp.createContact(firstName, theRest, member.email)
  if (member.homePhone)   {contact.addPhone("Home", member.homePhone)}
  if (member.mobile) {contact.addPhone("Mobile", member.mobile)}
  if (member.homeAddress) {contact.addAddress(ContactsApp.Field.HOME_ADDRESS, member.address)}
  
  setID(contact, member.id)
  
  contact.addToGroup(group)
  
  // add contact to system Contacts group or will not be able to manually edit contact
  contact.addToGroup(ContactsApp.getContactGroup("System Group: My Contacts"))

  log(["Added member to contacts", member])
}  


function updateContact(contact, member, coopGroup){
  var groups  
  var exCoopGroup = ContactsApp.getContactGroup("Ex members")

  // Id  
  if (!hasID(contact)) {
    setID(contact, member.id)
    
    // add to members group unless already in members group or in ex Members group
    groups = contact.getContactGroups()
//    if (!_.){
//      contact.addToGroup(coopGroup)
//    }

  } else {
    if (member.id !== getID(contact)){
      log(["ERROR updating contact", "Member id is different from the id already recorded in the contact", member, getID(contact) ])
      return
    }
  }
  
  // other fields
  updateMobile(contact, member.mobile)
  updateHomePhone(contact, member.homePhone)
  updateAddress(contact, member.address)
  updateAddress(contact, member.address)

  
                                 
  // ...Email...
  
}


function updateMobile(contact, mobile) {
  var phones = contact.getPhones('Mobile') 
  if (!mobile) {
    for (var p in phones) {
      log(['Contact Updated', 'Removed mobile number from', contact.getFullName(), phones[p].getPhoneNumber()])
      phones[p].deletePhoneField()
    }
    return
  }
  
  if (phones.length == 0) {
    contact.addPhone("Mobile", mobile)
    log(['Contact Updated', 'Added mobile number for', contact.getFullName(), mobile])
  } else {
    if (mobile !== phones[0].getPhoneNumber()) {
      log(['Contact Updated', 'Changed mobile number for', contact.getFullName(), 'from '+ phones[0].getPhoneNumber(), 'to '+ mobile])
      phones[0].setPhoneNumber(mobile).setLabel("Mobile")
    }
  }
}

function updateHomePhone(contact, homephone) {
  var phones = contact.getPhones("Home") 
  if (!homephone) {
    for (var p in phones) {
      log(['Contact Updated', 'Removed home phone number from', contact.getFullName(), phones[p].getPhoneNumber()])
      phones[p].deletePhoneField()
    }
    return
  }
  
  if (phones.length == 0) {
    contact.addPhone("Home", homephone)
    log(['Contact Updated', 'Added home phone number for', contact.getFullName(), homephone])
  } else {
    if (homephone !== phones[0].getPhoneNumber()) {
      log(['Contact Updated', 'Changed home phone number for', contact.getFullName(), 'from '+ phones[0].getPhoneNumber(), 'to '+ homephone])
      phones[0].setPhoneNumber(homephone).setLabel("Home")
    }
  }
}

function updateAddress(contact, address) {
  var addresses = contact.getAddresses('Home') 
  if (!address) {
    for (var p in addresses) {
      log(['Contact Updated', 'Removed address from', contact.getFullName(), addresses[p].getAddress()])
      addresses[p].deleteAddressField()
    }
    return
  }
  
  if (addresses.length == 0) {
    contact.addAddress("Home", address)
    log(['Contact Updated', 'Added address for', contact.getFullName(), address])
  } else {
    if (address !== addresses[0].getAddress()) {//address changed
      log(['Contact Updated', 'Changed address for', contact.getFullName(), 'from '+ addresses[0].getAddress(), 'to '+ address])
      addresses[0].setAddress(address).setLabel("Home")
    }
  }
}

function updateEmail(contact, email) {
  var emails = contact.getEmails() 
  if (!email) {
    for (var e in emails) {
      log(['Contact Updated', 'Removed email address from', contact.getFullName(), emails[e].getAddress()])
      emails[e].deleteEmailField()
    }
    return
  }
  
  if (emails.length == 0) {
    contact.addEmail("Other", email)
    log(['Contact Updated', 'Added email address for', contact.getFullName(), email])
  } else {
    if (email !== emails[0].getAddress()) {
      log(['Contact Updated', 'Changed email address for', contact.getFullName(), 'from '+ emails[0].getEmail(), 'to '+ email])
      emails[0].setEmailAddress(email).setLabel("Email")
    }
  }
}

function hasID(contact) {
  var fields = contact.getCustomFields()
  var fresh = isFRESH
  for (var i =0; i < fields.length; i++) {
    if (fresh && fields[i].getLabel() == 'Fresh ID' || !fresh && fields[i].getLabel() == 'Dry ID') {return true}
  }
  return false
}

function setID(contact, id) {
  if (isFRESH) {
    contact.addCustomField("Fresh ID", id)
  } else {
    contact.addCustomField("Dry ID", id)
  }
}

function getID(contact) {
  var fields = contact.getCustomFields("Fresh ID")
  if (isFRESH) {
    contact.getCustomFields("Fresh ID")
  } else {
    contact.getCustomFields("Dry ID")
  }
  
  if (fields.length == 0) {
    return ""
  } else {
    if (fields.length == 1) {
      return fields[0].getValue()
    } else {
      log(["ERROR contact has " + fields.length + " id fields", {name: contact.getFullName(), id: fields[0].getValue()}])
      return fields[0].getValue()
    }
  }
}

function getFreshContacts() {
  var contacts = ContactsApp.getContactGroup("Co-op members").getContacts()
  //return contacts.filter(function (i) {return hasID(i)})
}



function getBanking(id) {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var data = ss.getSheetByName("Banking").getDataRange().getValues()
  if (data[0][7] !==  "BIN") {return}
  var transactions = []
  for (var i = data.length-1; i > 0; i--) {
    if (data[i][7] == id) {
      transactions.push({date: data[i][0],
                         amount: data[i][1],
                         id: data[i][7]})
    }
  }
  return transactions// all entries for id
}


function sendStatement(member){
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var endNotes = "<br><br>Balance quoted is as shown on the Totals tab of the latest spreadsheet at the time of sending. " + 
                 "Recent payments may not yet have shown up in our account." 
  var html =  
    '<body>' + 
      '<h3 style="color: green;">Kapiti Co-op - Statement</h3>' +
      '<h2> Statement </h2><br />' +
        '<p> Greetings Earthling </p>' +
          endNotes +
    '</body>'

  var message = {to: "flourish.nz@gmail.com"     ,//   + ", " + recipients,
                 subject: "Fresh TEST statement " + ss.getName(),
                 htmlBody: html + "<br><br><a href='" + ss.getUrl() + "'>" + ss.getName() + "</a>"
                }
  
  MailApp.sendEmail(message)
}

//  tellJulie("<h1>Statement</h1><br>Your Fresh account is ((member.ProvisionalBalance < 0) ? "in debit." : "in credit.") +
//            " Your balance is $" + member.ProvisionalBalance  + "." +
//            
//            "<br>" +
//            //"Your most recent orders were ... <br><br>" +
//            "Your most recent payments were ... <br><br>"  +
//            "Payments made after ... have not yet been downloaded.<br> <br>" +
//            "<div foregroundcolor grey>This message has been automatically generated and is still experimental. </div>" +
//            
//            '<iframe id="forum_embed" ' +
//            'src="javascript:void(0)"' +
//            'scrolling="no"' +
//            'frameborder="0"' +
//            'width="900"' +
//            'height="700">' +
//            '</iframe>' +
//            '<script type="text/javascript">' +
//            "document.getElementById('forum_embed').src =" +
//    'https://groups.google.com/forum/embed/?place=forum/fresh-committee ' +
//   '&showsearch=true&showpopout=true&showtabs=false  '+
//   '&parenturl= + encodeURIComponent(window.location.href)' +
//  "</script>"
            
       

//function getActiveMembersNegative(){ 
//  var data = getLatestTotals()
//  return _.filter(data, function(m) {return true}) // && m.ProvisionalBalance <0)})
//}
//
//function getActiveMembersPositive(){ 
//  var data = getLatestTotals()
//  return _.filter(data, function(m) {return (m.Contingency > 0 && m.ProvisionalBalance >= 0)})
//}

function getLatestTotals(optID){
  var ID = optID || ""
  var ssObj = getLatestSS("Fresh Orders 20")  
  var totals = getTotals(ssObj)
  var re = /^\d{4}$/

  if (isValidId(ID)) {     
    return totals[ID]
  }
  else {
    return totals
  }
}


function getTotals(ssObj){//tested for Fresh only

// if (!ssObj){return};
  var ss = SpreadsheetApp.open(ssObj)
  var arr = []

  var sheet = ss.getSheetByName("Totals")
  var data = sheet.getDataRange().getValues()
  
// Locate, select and transpose the totals section
  var iSubtotal = ArrayLib.indexOf(data, 0, "Sub-total")
  var balances = data.slice(iSubtotal, iSubtotal+20)
  var transBalances = ArrayLib.transpose(balances)
      
// Locate headers
  var iContingency = ArrayLib.indexOf(balances, 0, "Contingency (wastage, errors, admin...)")
  var iAccounting = ArrayLib.indexOf(balances, 0, "Accounting and IT")
  var iMembership = ArrayLib.indexOf(balances, 0, "Membership")
  var iOrderTotal = ArrayLib.indexOf(balances, 0, "Order Total")
  
  var iNames = ArrayLib.indexOf(balances, 0, "Fresh Accounts")
  var iIDs = ArrayLib.indexOf(balances, 0, "Totals")
  
  var iPreviousBalance = ArrayLib.indexOf(balances, 0, "Previous Order")
  var iPreviousOrder = ArrayLib.indexOf(balances, 4, "Previous Order")
  var iPreviousPayments = ArrayLib.indexOf(balances, 4, "payments received between...")
  var iPreviousOtherCredits = ArrayLib.indexOf(balances, 4, "Other credits and debits")
  
  var iCurrentBalance = ArrayLib.indexOf(balances, 0, "Current Order")
  var iCurrentOrder = ArrayLib.indexOf(balances, 4, "Current order total (as above)")
  var iCurrentPayments = iCurrentOrder + 1


  var iProvisionalBalance = ArrayLib.indexOf(balances, 4, "Provisional Balance")
  
// initalise 
  var member = []
  var id = ""
  var previousBalanceDate = new Date(balances[iPreviousBalance, 4])
  var currentBalanceDate = new Date(balances[iCurrentBalance, 4])
  var provisionalBalanceDate = new Date(balances[iCurrentPayments + 1, 4])
  var totals = {}

// step through the members
  
  for(var i = 5; i < balances[6].length; i++){   // 5 is the first member column - get this from somewhere?

    member = transBalances[i];
    id = member[iIDs].toString()
    
    if (isValidId(id)) {
      totals = {
        ID: id,
        Name: member[iNames],
        
        Contingency: member[iContingency],
        Accounting: member[iAccounting],
        Membership: member[iMembership],
        OrderTotal: member[iOrderTotal],
        
        PreviousBalance: member[iPreviousBalance],
        PreviousOrder: member[iPreviousOrder],
        PreviousPayments: member[iPreviousPayments],
        PreviousOtherCredits: member[iPreviousOtherCredits],
        
        CurrentBalance: member[iCurrentBalance],
        CurrentOrder: member[iCurrentOrder],
        CurrentPayments: member[iCurrentPayments],
        
        ProvisionalBalance: member[iProvisionalBalance]
      }

      arr[id] = totals
    }
  }
  return arr
}



function isValidId(id) {
  return /^\d{4}$/.test(id)
}


