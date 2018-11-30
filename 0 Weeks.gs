function runAllFormattingFunctions_upcomingWeek_(doc) { // this week

  doc = doc || DocumentApp.getActiveDocument();
  format_upcoming_font();//sets normal font and spacing
  format_upcoming_highlightEventName();
  format_boldBetweenSquareBrackets();
  format_unHighlightStaffNames();
  format_removeEmptyParagraphs_();
  format_titleAndSubtitle();//and set para style so it appears on the outline
  return; 
  
  // Private Functions
  // -----------------
  
  function format_titleAndSubtitle(){
  
    var body = doc.getBody();
  
    var re_ordinalSunday = /^ *(?:First|Second|Third|Fourth|Fifth) *Sunday *of *the *month *$/i;
    var re_ordinalSunday_string = escapeGasRegExString(re_ordinalSunday);
    var ordinalElement = body.findText(re_ordinalSunday_string);
    if(ordinalElement){
      var ordinalParagraph = ordinalElement.getElement().getParent();
      ordinalParagraph.asParagraph().setHeading(DocumentApp.ParagraphHeading.HEADING2);
      ordinalParagraph.setAttributes(config.announcements.format.subtitle)
    }
  
    var titleElement = body.findText(config.announcements.sundayPagePattern);
    
    if(titleElement){
      var titleParagraph = titleElement.getElement().getParent();
      titleParagraph.asParagraph().setHeading(DocumentApp.ParagraphHeading.TITLE)
      titleParagraph.setAttributes(config.announcements.format.title)
    }
    
  }
  
  function format_upcoming_font() {
    var body = doc.getBody();
    body.editAsText().setFontFamily('Lato');
    body.editAsText().setFontSize(9);
    body.editAsText().setForegroundColor('#58585a');
    var p = body.getParagraphs();
    for (i = 0; i < p.length; i++) {
      p[i].setLineSpacing(1.5).setSpacingAfter(0);
    }
  }
  
  function format_upcoming_highlightEventName() {//and set bold
    var body = doc.getBody();
    var para = body.getParagraphs();
    var startTag = "[\[]";
    var endTag = "[|]";
    var highlightStyle = {};
    highlightStyle[DocumentApp.Attribute.BACKGROUND_COLOR] = '#FCFC00';
    for (var i in para) {
      var from = para[i].findText(startTag);
      var to = para[i].findText(endTag, from);
      if ((to != null && from != null) && ((to.getStartOffset()) - (from.getStartOffset() + startTag.length) > 0)) {
        para[i].editAsText().setBold(from.getStartOffset() + 2, to.getStartOffset() - 2, true);
        para[i].editAsText().setBackgroundColor(from.getStartOffset() + 2, to.getStartOffset() - 2, "#FCFC00");
      }
    }
  }
  
  function format_boldBetweenSquareBrackets() {
    var body = doc.getBody();
    var para = body.getParagraphs();
    var startTag = "[\[]";
    var endTag = "[\]]";
    var i = 0;
    for (i in para) {
      var from = para[i].findText(startTag);
      var to = para[i].findText(endTag, from);
      if ((to != null && from != null) && ((to.getStartOffset()) - (from.getStartOffset() + startTag.length) > 0)) {
        para[i].editAsText().setBold(from.getStartOffset(), to.getStartOffset(), true);
      }
    }
  }
  
  function format_unHighlightStaffNames() {
    var body = doc.getBody();
    var para = body.getParagraphs();
    var startTag = "[;]";
    var endTag = "[\]]";
    var highlightStyle = {};
    highlightStyle[DocumentApp.Attribute.BACKGROUND_COLOR] = '#FFFFFF';
    var i = 0;
    for (i in para) {
      var from = para[i].findText(startTag);
      var to = para[i].findText(endTag, from);
      if ((to != null && from != null) && ((to.getStartOffset()) - (from.getStartOffset() + startTag.length) > 0)) {
        para[i].editAsText().setBold(from.getStartOffset() + 1, to.getStartOffset() - 1, true);
        para[i].editAsText().setBackgroundColor(from.getStartOffset() + 1, to.getStartOffset() - 1, "#FFFFFF");
      }
    }
  }
  
} // runAllFormattingFunctions_upcomingWeek_()
