function runAllFormattingFunctions_oneWeekOut_(doc) {
  doc = doc || DocumentApp.openById(Config.get('ANNOUNCEMENTS_1WEEK_SUNDAY_ID'));
  paragraphs = doc.getBody().getParagraphs();
  formatFont_oneWeek_(doc);
  format_removeEmptyParagraphs_(paragraphs);
}

function formatFont_oneWeek_(doc) {
  doc = doc || DocumentApp.openById(Config.get('ANNOUNCEMENTS_1WEEK_SUNDAY_ID'))
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