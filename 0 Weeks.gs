function runAllFormattingFunctions_upcomingWeek_(doc) { // this week

  doc = doc || DocumentApp.getActiveDocument();
  var body = doc.getBody();
  var paragraphs = body.getParagraphs();
    
  format_upcoming_font();//sets normal font and spacing
  format_upcoming_highlightEventName();
  format_boldBetweenSquareBrackets();
  format_unHighlightStaffNames();
  format_removeEmptyParagraphs_(paragraphs);
  format_titleAndSubtitle();//and set para style so it appears on the outline
  return; 
  
  // Private Functions
  // -----------------
  
  function format_titleAndSubtitle(){
  
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
  
    body.editAsText().setFontFamily('Lato');
    body.editAsText().setFontSize(9);
    body.editAsText().setForegroundColor('#58585a');
    
    for (i = 0; i < paragraphs.length; i++) {
      paragraphs[i].setLineSpacing(1.5).setSpacingAfter(0);
    }
  }
  
  function format_upcoming_highlightEventName() {//and set bold
  
    var startTag = "[\[]";
    var endTag = "[|]";
    var highlightStyle = {};
    highlightStyle[DocumentApp.Attribute.BACKGROUND_COLOR] = '#FCFC00';
    for (var i in paragraphs) {
      var from = paragraphs[i].findText(startTag);
      var to = paragraphs[i].findText(endTag, from);
      if ((to != null && from != null) && ((to.getStartOffset()) - (from.getStartOffset() + startTag.length) > 0)) {
        paragraphs[i].editAsText().setBold(from.getStartOffset() + 2, to.getStartOffset() - 2, true);
        paragraphs[i].editAsText().setBackgroundColor(from.getStartOffset() + 2, to.getStartOffset() - 2, "#FCFC00");
      }
    }
  }
  
  function format_boldBetweenSquareBrackets() {
  
    var startTag = "[\[]";
    var endTag = "[\]]";
    var i = 0;
    for (i in paragraphs) {
      var from = paragraphs[i].findText(startTag);
      var to = paragraphs[i].findText(endTag, from);
      if ((to != null && from != null) && ((to.getStartOffset()) - (from.getStartOffset() + startTag.length) > 0)) {
        paragraphs[i].editAsText().setBold(from.getStartOffset(), to.getStartOffset(), true);
      }
    }
  }
  
  function format_unHighlightStaffNames() {
  
    var startTag = "[;]";
    var endTag = "[\]]";
    var highlightStyle = {};
    highlightStyle[DocumentApp.Attribute.BACKGROUND_COLOR] = '#FFFFFF';
    var i = 0;
    for (i in paragraphs) {
      var from = paragraphs[i].findText(startTag);
      var to = paragraphs[i].findText(endTag, from);
      if ((to != null && from != null) && ((to.getStartOffset()) - (from.getStartOffset() + startTag.length) > 0)) {
        paragraphs[i].editAsText().setBold(from.getStartOffset() + 1, to.getStartOffset() - 1, true);
        paragraphs[i].editAsText().setBackgroundColor(from.getStartOffset() + 1, to.getStartOffset() - 1, "#FFFFFF");
      }
    }
  }
  
} // runAllFormattingFunctions_upcomingWeek_()
