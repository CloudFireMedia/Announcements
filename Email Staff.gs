function emailStaff_() {

  var staffDocId = Config.get('STAFF_DATA_GSHEET_ID');
  var staffDoc = SpreadsheetApp.openById(staffDocId);
  var html = '<!DOCTYPE html><head><title>Select Names</title><script src="//ajax.googleapis.com/ajax/libs/jquery/1.9.1/jquery.min.js"></script></head><body>';
  html += '<script> function sendEmailsSuccess(){ google.script.host.close(); } function sendEmails(){ var ea=[]; var ei=0; $(\'select[name=e1] option:selected\').each(function(){ ea[ei]=$(this).attr(\'value\'); ei++ }); $(\'select[name=e2] option:selected\').each(function(){ ea[ei]=$(this).attr(\'value\'); ei++ }); google.script.run.withSuccessHandler(sendEmailsSuccess).emailStaff_submit(JSON.stringify(ea)); } </script>';
  var sheet = staffDoc.getActiveSheet();
  var range = sheet.getRange(3, 1, sheet.getLastRow() - 2, 2);
  var ra = range.getDisplayValues();
  
  var hselect = '<option value="-1"> </option>';
  
  for (var ri = 0; ri < ra.length; ri++) {
    hselect += '<option value="' + ri + '"> ' + ra[ri][0] + ' ' + ra[ri][1] + '</option>';
  }
  
  html += 'Executive Assistant: <select name="e1">' + hselect + '</select><br/>';
  html += 'Media Director: <select name="e2">' + hselect + '</select><br/>';
  html += '<input type="button" value="Send E-mails" onClick="javascript:{sendEmails();}"/></body></html>';
  var output = HtmlService.createHtmlOutput(html);
  output.setTitle('Select Users to send E-mails');
  DocumentApp.getUi().showSidebar(output);
  
}

function emailStaff_submit_(jea) {

  var thisWeekDocId = Config.get('ANNOUNCEMENTS_0WEEKS_SUNDAY_ID');
  var thiwWeekDocUrl = DocumentApp.openById(thisWeekDocId).getUrl();
  
  var folderId = Config.get('SLIDES_PARENT_FOLDER_ID');
  var folderUrl = DriveApp.getFolderById(folderId).getUrl();
  
  var ea = JSON.parse(jea);
  var staffDocId = Config.get('STAFF_DATA_GSHEET_ID');
  var doc = SpreadsheetApp.openById(staffDocId);
  var sheet = doc.getActiveSheet();
  var range = sheet.getRange(3, 9, sheet.getLastRow() - 2, 1);
  var ra = range.getDisplayValues();
  var range2 = sheet.getRange(3, 1, sheet.getLastRow() - 2, 2);
  var ra2 = range2.getDisplayValues();
  
  var s = '';
  for (var ei = 0; ei < ea.length; ei++) {
    if (ea[ei] != '-1') {
      var i = parseInt(ea[ei]);
      s += 'VALUE=' + ra[i][0] + ';';
    }
  }
  
  //Executive Assistant
  if (ea[0] != '-1') {
    var subject = 'This Sunday\'s Announcements have just been updated';
    var body = ra2[ea[0]][0] + ' ' + ra2[ea[0]][1] + ', FYI this Sunday\'s announcements have just been updated, due to a last minute addition.';
    var htmlBody = ra2[ea[0]][0] + ' ' + ra2[ea[0]][1] + ', FYI <a href="' + thiwWeekDocUrl + '">this Sunday\'s announcements</a> have just been updated, due to a last minute addition.';
    var to = ra[ea[0]][0];
    MailApp.sendEmail({
      to: to,
      subject: subject,
      body: body,
      htmlBody: htmlBody
    });
  }
  
  if (ea[1] != '-1') {
    var subject = 'This Sunday\'s Announcements have just been updated';
    var body = ra2[ea[1]][0] + ' ' + ra2[ea[1]][1] + ', FYI this Sunday\'s Service Slides have just been updated, due to a last minute addition.';
    var htmlBody = ra2[ea[1]][0] + ' ' + ra2[ea[1]][1] + ', FYI <a href="' + folderUrl + '">this Sunday\'s Service Slides</a> have just been updated, due to a last minute addition.';
    var to = ra[ea[1]][0];
    MailApp.sendEmail({
      to: to,
      subject: subject,
      body: body,
      htmlBody: htmlBody
    });
  }
  //MailApp.sendEmail("w2kzx80@gmail.com", "TPS reports", s);
}
