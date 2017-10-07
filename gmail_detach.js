/*
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Delete Gmail attachments while retaining a copy of the original email
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Heavily inspired by Shunmugha Sundaram's script:
http://techawakening.org/?p=1842
*/
var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
var sheet       = spreadsheet.getSheetByName('Emails');
var mib         = 1048576;
var firstRow    = 9;

// adding menu
function onOpen() {
  var menuEntries = [
    {name: "Search emails",       functionName: "searchEmails"},
    {name: "Mark all for detach", functionName: "markAll"},
    {name: "Unmark all",          functionName: "unmarkAll"},
    {name: "Delete attachments",  functionName: "processEmails"}
  ];
  spreadsheet.addMenu("GmailDetach", menuEntries);
}

function searchEmails() {
  var row = firstRow;
  var ct  = 0;
  clearSheet();
  sheet.getRange(firstRow,1).activate()
  getThreads().forEach(function(t) {
    t.getMessages().forEach(function(msg) {
      var att = msg.getAttachments();
      if (att.length > 0) {
        sheet.getRange(row,2).setValue(msg.getId());
        sheet.getRange(row,3).setValue(msg.getFrom());
        sheet.getRange(row,4).setValue(msg.getSubject());
        sheet.getRange(row,5).setValue(msg.getDate());
        sheet.getRange(row,6).setValue(att.length)
        sheet.getRange(row,6).setNote(attNames(att));
        sheet.getRange(row,7).setValue(attSize(att));
        row++;
      }
    })
    monitorSearch(ct++)
  })
  sheet.getRange(7,1).setValue('Fetched emails');
}

function processEmails() {
  var s = 'Messages marked with an ‘x’ will be moved to the Gmail trash; ' +
          'their attachments will be backed up on your Drive ' +
          'and a copy of the original message will be sent to your address. ' +
          'Continue?'
  if (Browser.msgBox('Heads up!', s, Browser.Buttons.YES_NO) != 'yes') return;
  for (var row = firstRow; row <= sheet.getLastRow(); row++) {
    if (toDel(row)) {
      var msg = theMsgAtRow(row);
      msg.getAttachments().forEach(function(att) {
        theFolderFor(msg).createFile(att);
      });
      sendEmailFor(msg);
      msg.moveToTrash();
      resetRow(row);
    }
  }
}

function monitorSearch(ct) {
  return sheet.getRange(7,1).setValue('Fetching... ' +  (maxThreads() - ct))
}

function markAll() {
  markem('x')
}

function unmarkAll() {
  markem('')
}

function markem(x) {
  var nr = sheet.getLastRow() - firstRow + 1;
  var arr = [];
  for (var i = 0; i < nr; i++) {
    arr.push([x])
  }
  sheet.getRange(firstRow,1,nr).setValues(arr)
}

function getThreads() {
  var s = 'has:attachment'
  if (threadSize())   { s += ' larger:' + threadSize()   }
  if (beforeDate()) { s += ' before:' + beforeDate().toISOString().substr(0, 10) }
  if (afterDate())  { s += ' after:'  + afterDate().toISOString().substr(0, 10)  }
  spreadsheet.toast('Searching emails with the following query: ' + s)
  return GmailApp.search(s, 0, maxThreads())
}

function attNames(attachments) {
  var s = '';
  attachments.forEach(function(a) {
    s += "➜ " + a.getName() + "\n";
  })
  return s
}

function attSize(attachments) {
  var s = 0;
  attachments.forEach(function(a) {
    s += a.getSize();
  })
  return s/mib
}

function sendEmailFor(msg) {
  GmailApp.sendEmail(
    Session.getActiveUser().getUserLoginId(),
    msg.getSubject(),
    '',
    { htmlBody: mailBody(msg) }
  );
}

function theFolderFor(msg) {
  var rf = ensureFolder(DriveApp.getRootFolder(), backupFolderName());
  var df = ensureFolder(rf, dateString(msg));
  var fs = msg.getSubject() + ' (' + msg.getId() + ')';
  return ensureFolder(df, fs)
}

function ensureFolder(parent, name) {
  var ef = parent.getFoldersByName(name);
  if (ef.hasNext()) {
    return ef.next()
  } else {
    return parent.createFolder(name);
  }
}

function dateString(msg) {
  return msg.getDate().toISOString().substr(0, 10);
}

function mailBody(msg) {
  var br = '<br/>';
  var body =
    '<b>GmailDetach saved the following attachments in the folder:</b>' + br +
    attPath(msg) + br + br + attList(msg)                               + br +
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'        + br +
    '<b>The following original email was sent to the Trash:</b>'        + br + br +
    'From: '    + msg.getFrom()                                         + br +
    'Subject: ' + msg.getSubject()                                      + br +
    'Date: '    + msg.getDate()                                         + br +
    'To: '      + msg.getTo()                                           + br +
    'CC: '      + msg.getCc()                                           + br +
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'        + br + br +
    msg.getBody();
  return body;
}

function attPath(msg) {
  return backupFolderName() +'/'+ dateString(msg) +'/'+ msg.getSubject();
}

function attList(msg) {
  var s = '';
  msg.getAttachments().forEach(function(a) {
    s += '➜ ' + a.getName() + '<br/>';
  })
  return s
}

function attFolder() {
  var r = DriveApp.getRootFolder();
  var fs = r.getFoldersByName('Gmail-Attachments');
  if (fs.hasNext()) {
    var fol = fs.next();
  } else {
    var fol = DriveApp.createFolder('Gmail-Attachments');
  }
  return fol
}

// spreadsheet-aware functions

function toDel(row) {
  var s = sheet.getRange(row, 1).getValue();
  return s.toLowerCase() == 'x' && msgIdAtRow(row)
}

function clearSheet() {
  var r = sheet.getRange(firstRow, 1, sheet.getLastRow(), 8);
  r.clearContent();
  r.setFontColor('black');
  r.clearNote();
}

function theMsgAtRow(row) {
  return GmailApp.getMessageById(msgIdAtRow(row))
}

function resetRow(row) {
  sheet.getRange(row,1,1,2).clearContent();
  sheet.getRange(row, 8).setValue('OK');
  sheet.getRange(row,3,1,5).setFontColor('gray')
}

// READING PREFERENCES
function maxThreads() {
  var nt = spreadsheet.getRangeByName('nthreads').getValue();
  if (nt) {
    return parseInt(nt)
  } else {
    return 10
  }
}
function threadSize() {
  return spreadsheet.getRangeByName('thresize').getValue()
}
function backupFolderName() {
  return spreadsheet.getRangeByName('backupfol').getValue()
}

function afterDate() {
  return spreadsheet.getRangeByName('after').getValue();
}

function beforeDate() {
  return spreadsheet.getRangeByName('before').getValue();
}

function msgIdAtRow(row) {
  return sheet.getRange(row, 2).getValue();
}
