function runAllFormattingFunctions_oneWeekOut(doc) {
  doc = doc || DocumentApp.openById(config.files.announcements.oneWeek)
  formatFont_oneWeek(doc);
//  format_removeEmptyParagraphs_oneWeek(doc);
  format_removeEmptyParagraphs(doc);
}

function formatFont_oneWeek(doc) {
  doc = doc || DocumentApp.openById(config.files.announcements.oneWeek);//var doc = DocumentApp.openById(id)
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

//function format_removeEmptyParagraphs_oneWeek(doc) { ///rewrite based on the upcoming sunday version
//  doc = doc || DocumentApp.openById(config.files.announcements.oneWeek);
//  var body = doc.getBody();
//  var paras = body.getParagraphs(); 
//  for (var i = 0; i < paras.length; i++) { 
//    if (paras[i].getText() === ""){
//      if (paras[i].findElement(DocumentApp.ElementType.HORIZONTAL_RULE,n‌ull) === null)
//        if (paras[i].findElement(DocumentApp.ElementType.PAGE_BREAK,n‌ull) === null)
//          if( ! paras[i].isAtDocumentEnd())
//            paras[i].removeFromParent();
//    } 
//  } 
//} 

