// CONTACTS
// adding and removing contacts of new and ex members 
// can only add contacts to its own contacts list - so must be run by kapitidrycoop or kapitifresh.co.op
// may be able to use triggers to make this happen from other accounts
//
// 0.1 LIVE for dry (I think) using addMembertoContacts

function testMC(){
//  //var members = getMembers()
//  var member = {name: "my test", id: "9998", email: "about@example.com", homePhone: 432432,               address: "321 Lets Drive"}
//  addMemberToContacts(member)
//
//  var member = {name: "My Test", id: "9998", email: "about@example.com", mobile: "123", address: "321 Lets Drive"}

//  var members = getMembers()
//  for (var i = 0; i<members.length; i++){
//    if (members[i].firstName == "Fern") {addMemberToContacts(members[i])}
//  }
}


function updateMember(member) {
  // to be auto-run from co-op account onEdit of Members tab
}


/**
 * Adds/Updates Co-op contact
 *
 * @param {object} member the member information eg {name: "Fred", id: 2102,...}
 * @return 
 */
function addMemberToContacts(member) {
  var coopGroup = ContactsApp.getContactGroup("Co-op members")

  if (!coopGroup){
    log(["Cannot access Contact Group from this account", member])
    return
  }
  
  var contacts = ContactsApp.getContactsByName(member.name)
  if (contacts.length == 0) {
    addContact(member, coopGroup)
  } else if (contacts.length == 1){
    updateContact(contacts[0], member, coopGroup)
  } else {//... fix... handle this better - alert - etc
    log(["Multiple contacts exist with this name. Not updated", member])
    SpreadsheetApp.getUi().alert('Multiple contacts exist with this name. Not updated.\n' + member)
  }
}

function removeFromCurrentContacts(member) {
  //  Remove contact from current list - has to be run by coop account
  if (isFRESH){
    var coopGroup = ContactsApp.getContactGroup("Co-op members")
    var exGroup = ContactsApp.getContactGroup("Ex members")
    var contacts = ContactsApp.getContactsByName(member.name)
    if (contacts.length == 0) {
      log(["Contact not found", member.name])
    } else {
      for (var i in contacts) {
        exGroup.addContact(contacts[i])
        coopGroup.removeContact(contacts[i])
        log(["Moved contact from Co-op Members group to Ex Members group", contact[i].name])
      }
    }
  } 
  else {
    
  }

}


function addContact(member, group){
  if (isFRESH){
    var firstName = member.name.split(" ")[0]
    var theRest = member.name.substring(firstName.length+1, member.name.length)
    var contact = ContactsApp.createContact(firstName, theRest, member.email)
    } else {
      var contact = ContactsApp.createContact(member.firstName, member.lastName, member.email.trim())
    }
  
  if (member.homePhone)   {contact.addPhone("Home", member.homePhone)}
  if (member.mobile) {contact.addPhone("Mobile", member.mobile)}
  if (member.otherPhone) {contact.addPhone("Other", member.otherPhone)}
  if (member.homeAddress) {contact.addAddress(ContactsApp.Field.HOME_ADDRESS, member.homeAddress)}
  
  setID(contact, member.id)
  
  // add contact to group; also add to system Contacts group or will not be able to manually edit contact
  contact.addToGroup(group)
  contact.addToGroup(ContactsApp.getContactGroup("System Group: My Contacts"))
  log(["Added member to contacts", member])
}  


function updateContact(contact, member, coopGroup){
  // add to members group unless already in members group or in ex Members group
  var groups = contact.getContactGroups()
  if (groups.indexOf(coopGroup) == -1 && groups.indexOf(ContactsApp.getContactGroup("Ex members")) == -1){
    log(['Member updated', "Updating contacts", member.name, member.id])
    contact.addToGroup(coopGroup)
  }
  
  // other fields
  if (updateId_(contact, member.id)) {
    updateMobile_(contact, member.mobile)
    updateHomePhone_(contact, member.homePhone)
    updateOtherPhone_(contact, member.otherPhone)
    updateAddress_(contact, member.homeAddress)
    updateEmail_(contact, member.email) 
  }
}


function updateId_(contact, id){//... problem here as the IDs never match - type problem
  // Id  - add id if missing - quit if id exists but doesn't match
  if (!isValidId(id)){
    log(["ERROR UPDATING CONTACT", "Invalid id supplied", id])
    return -1
  }
  if (!hasID(contact)) {
    setID(contact, id)
  } else {
    if (id !== getID(contact)){
      log(["ERROR UPDATING CONTACT", "Member id is different from the id already recorded in the contact", 'New: ' +id, getID(contact) ])
      return -1
    }
  }
  return 1
}

function updateMobile_(contact, mobile) {
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

function updateHomePhone_(contact, homephone) {
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

function updateOtherPhone_(contact, otherphone) {
  var phones = contact.getPhones("Other") 
  if (!otherphone) {
    for (var p in phones) {
      log(['Contact Updated', 'Removed other phone number from', contact.getFullName(), phones[p].getPhoneNumber()])
      phones[p].deletePhoneField()
    }
    return
  }
  
  if (phones.length == 0) {
    contact.addPhone("Other", otherphone)
    log(['Contact Updated', 'Added other phone number for', contact.getFullName(), otherphone])
  } else {
    if (otherphone !== phones[0].getPhoneNumber()) {
      log(['Contact Updated', 'Changed other phone number for', contact.getFullName(), 'from '+ phones[0].getPhoneNumber(), 'to '+ otherphone])
      phones[0].setPhoneNumber(otherphone).setLabel("Other")
    }
  }
}

function updateAddress_(contact, address) {//... not convinced about this one
  log('called address' + address)
  var addresses = contact.getAddresses('Home')
  if (!address) {
    for (var p in addresses) {
      log(['Contact Updated', 'Removed address from', contact.getFullName(), addresses[p].getAddress()])
      addresses[p].deleteAddressField()
    }
    return
  }
  
  if (addresses.length == 0) {
    log(['Contact Updated', 'Added address for', contact.getFullName(), address])
    contact.addAddress("Home", address)
  } else {
    if (address !== addresses[0].getAddress()) {//address changed
      log(['Contact Updated', 'Changed address for', contact.getFullName(), 'from '+ addresses[0].getAddress(), 'to '+ address])
      addresses[0].setAddress(address).setLabel("Home")
    } else {
      log('Huh?')
    }
  }
}

function updateEmail_(contact, email) {
  var emails = contact.getEmails() 
  if (!email) {
    for (var e in emails) {
      var msg = ['Contact Updated', 'Removed email address', contact.getFullName(), emails[e].getAddress()]
      emails[e].deleteEmailField()
      log(msg)
    }
    return
  }
  
  if (emails.length == 0) {
    contact.addEmail("Other", email)
    log(['Contact Updated', 'Added email address for', contact.getFullName(), email])
  } else {
    if (email !== emails[0].getAddress()) {
      var msg = ['Contact Updated', 'Changed email address', contact.getFullName(), 'from '+ emails[0].getAddress(), 'to '+ email]
      emails[0].setAddress(email).setLabel("Email")
      log(msg)
    }
  }
}

function hasID(contact) {
  var fields = contact.getCustomFields()
  for (var i =0; i < fields.length; i++) {
    if (isFresh() && fields[i].getLabel() == 'Fresh ID' || isDry() && fields[i].getLabel() == 'Dry ID') {return true}
  }
  return false
}

function setID(contact, id) {
  if (isFresh()) {
    contact.addCustomField("Fresh ID", id)
  } else {
    contact.addCustomField("Dry ID", id)
  }
}

function getID(contact) {
  var fields = contact.getCustomFields("Fresh ID")
  if (isFresh()) {
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
  var contacts = ContactsApp.getContactGroup("Co-op members local").getContacts()
  //return contacts.filter(function (i) {return hasID(i)})
}


