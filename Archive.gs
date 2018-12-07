function formatFont_archive_() {
  var doc = DocumentApp.openById(Config.get('ANNOUNCEMENTS_ARCHIVE_ID'));
  var body = doc.getBody();
  body.editAsText().setFontFamily('Lato');
  body.editAsText().setFontSize(9);
  body.editAsText().setForegroundColor('#58585a');
  var p = body.getParagraphs();
  for(i=0;i<p.length; i++){
    p[i].setLineSpacing(1.5).setSpacingAfter(0);
    p[i].setBold(false);
    p[i].setBackgroundColor('#FFFFFF')
  }
}

