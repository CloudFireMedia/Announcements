function format_master_TRIGGERED(){format_master_()}

function format_master_(doc) {

  if (doc === undefined) {
  
    doc = DocumentApp.getActiveDocument();
  
    if (doc === null) {
      var masterDocId = Config.get('ANNOUNCEMENTS_MASTER_SUNDAY_ID');
      doc = DocumentApp.openById(masterDocId);
    }
  }

  var body = doc.getBody();
  var paragraphs = body.getParagraphs();
  
  format_doubleSpaceToSingle(doc);
  format_removeEmptyParagraphs_(paragraphs);
  format_master_fixPageBreaksAndHeadings_(doc);
  format_master_formatEventsFuture(doc);
  format_master_formatEventsPast(doc);
  
  // Private Functions
  // -----------------
    
  function format_doubleSpaceToSingle(doc) {
    body.replaceText("[ ]{2,}", " "); 
  }
  
  function format_master_formatEventsFuture(doc) {
    
    //find the announcements page for the next Sunday from today
    var nextSunday = getUpcomingSunday();
    //  var nextSunday = getUpcomingSunday(new Date('2018-04-01'));///DEV
    var shortDate = fDate(nextSunday, "'\\[' MM'\\.'dd '] Sunday Announcements'");
    var nextSundayRange = body.findText(shortDate);//this gets to the page title
    if( ! nextSundayRange) throw 'Could not find page "'+(shortDate.replace(/\\/g,''))+'" for this Sunday';
    var endOfNextSunday = body.findText(config.announcements.sundayPagePattern, nextSundayRange);//this gets to the NEXTpage title, need to back up one paragraph
    if( ! endOfNextSunday) throw 'Could not find page after "'+(shortDate.replace(/\\/g,''))+'"';
    var endOfNextSundayParagraphs = endOfNextSunday.getElement().getParent().getPreviousSibling();
    var end = body.getChildIndex(endOfNextSundayParagraphs);
    //  doc.setCursor(doc.newPosition(nextSundayElement, 0));
    
    //find first page with announcements
    var latestSundayRange = body.findText(config.announcements.sundayPagePattern);//the first instance of a Sunday page
    if( ! latestSundayRange) throw 'Could not find any announcement pages';//this should never happen
    var latestSundayElement = latestSundayRange.getElement()
    var latestSundayParagraph = latestSundayElement.getParent();
    var start = body.getChildIndex(latestSundayParagraph);
    //  doc.setCursor(doc.newPosition(latestSundayElement, 0));
    
    var eventTitleSearch = '^\\[.*?]';//find any string starting with [text in brackets]
    
    for(var i=start; i<=end; i++) {
      paragraphs[i].setAttributes(config.announcements.format.future);
      var eventTitleRange = paragraphs[i].findText(eventTitleSearch);
      if( ! eventTitleRange) continue;
      if(eventTitleRange.getElement().asText().getText().match(config.announcements.sundayPageRegEx)) continue;//don't mess with page titles
      eventTitleRange.getElement().asText().setAttributes(eventTitleRange.getStartOffset(), eventTitleRange.getEndOffsetInclusive(), config.announcements.format.futureTitle);
    }
  }
  
  function format_master_formatEventsPast(doc) {
  
    //find the announcements page for the Sunday following today
    var nextSunday = getUpcomingSunday();
    //  var nextSunday = getUpcomingSunday(new Date('2018-04-01'));///DEV
    var lastSunday = dateAdd(nextSunday, 'week', -1);
    var shortDate = fDate(lastSunday, "'\\[' MM'\\.'dd '] Sunday Announcements'");
    var lastSundayRange = body.findText(shortDate);//this gets to the page title
    if( ! lastSundayRange) throw 'Could not find page "'+(shortDate.replace(/\\/g,''))+'" for last Sunday';
    
    var start = body.getChildIndex(lastSundayRange.getElement().getParent());
    var end = paragraphs.length-1;
    for(var i=start; i<=end; i++) {
      if(paragraphs[i].getText().match(config.announcements.sundayPageRegEx)) continue;//don't mess with page titles
      paragraphs[i].setAttributes(config.announcements.format.past);
    }
    
    //this section only affects the [text in brackets], but we want the entire paragraph formatted for this one
    //  var eventTitleSearch = '^\\[.*?]';//find any string starting with [text in brackets]
    //  var substringRange = lastSundayRange;//start from last Sunday and continue to end of doc
    //  while( substringRange = body.findText(eventTitleSearch, substringRange) ){
    //    //doc.setCursor(doc.newPosition(substringRange.getElement(), 0));return
    //    var index = body.getChildIndex(substringRange.getElement().getParent());
    //    if(substringRange.getElement().asText().getText().match(config.announcements.sundayPageRegEx)) continue;//don't mess with page titles
    //    //    log(substringRange.getElement().asText().getText());
    //    substringRange.getElement().asText().setAttributes(substringRange.getStartOffset(), substringRange.getEndOffsetInclusive(), config.announcements.format.past)
    //  }
  }
  
} // format_master_()

function format_master_fixPageBreaksAndHeadings_(doc){

  var body = doc.getBody();
  
  ///note: consider adding a check for heading2 (nth Sunday) to ensure there is a paragraph between it and the pagebreak
  
  //remove all hard page breaks
  var elementType = DocumentApp.ElementType.PAGE_BREAK;
  var searchResult = null;
  //remove all pagebreaks
  while(searchResult = body.findElement(elementType, searchResult)){
    //var searchResult = body.findElement(elementType);
    var pagebreak = searchResult.getElement();
    pagebreak.removeFromParent();
    //searchResult is no longer valid after removing it (apparently) so grab it again then loop - at least, it only does one and stops without doing so
    searchResult = body.findElement(elementType, searchResult);
  }
  
  //find page titles
  var pageTitleSearchPattern = '\\[ \\d{2}\\.\\d{2} ] Sunday Announcements|\\[ RECURRING CONTENT ]|\\[.*?END.*?]';//'[ 03.11 ] Sunday Announcements' or '[ END OF DOC ]' or '[ RECURRING CONTENT ]' -- Note you must double-escape reserved characters since this is a text pattern not a regex
  var re_ordinalSunday = /^ *(?:First|Second|Third|Fourth|Fifth) Sunday of the month *$/i;
  var pageTitle = null;
  
  //add pagebreaks and format
  while(pageTitle = body.findText(pageTitleSearchPattern, pageTitle)){
    //var pageTitle = body.findText(searchPattern)
    var parent = pageTitle.getElement().getParent();
    var previousSibling = parent.getPreviousSibling();
    var nextSibling = parent.getNextSibling();
    if(previousSibling && previousSibling.getType() == DocumentApp.ElementType.PARAGRAPH){//it really should always be a paragraph but just in case...
      previousSibling.asParagraph().appendPageBreak();//insert a pagebreak in previous paragraph, before current
      parent.asParagraph().setHeading(DocumentApp.ParagraphHeading.HEADING1);//set its style to heading1
      parent.setAttributes(config.announcements.format.heading1)
      
      //see if there is an ordinal subtitle.  If so, set it to heading2      
      if(nextSibling && nextSibling.asParagraph().getText().match(re_ordinalSunday)){
        //this should be the hrule, not the nth Sunday text but we have to check because humans
        nextSibling.asParagraph().setHeading(DocumentApp.ParagraphHeading.HEADING2);
        nextSibling.setAttributes(config.announcements.format.heading2);
      }else{
        //now if it wasn't in the last one, which it should not be, then we check the next next one
        //this should match every pageTitle except RECURRING CONTENT and END OF DOCUMENT
        var nextNextSibling = nextSibling.getNextSibling();
        if(nextNextSibling && nextNextSibling.asParagraph().getText().match(re_ordinalSunday)){
          nextNextSibling.asParagraph().setHeading(DocumentApp.ParagraphHeading.HEADING2);
          nextNextSibling.setAttributes(config.announcements.format.heading2);
        }
      }
    }
  }
}

function format_removeEmptyParagraphs_(paragraphs) {

  if (paragraphs === undefined) {
    var doc = doc || DocumentApp.getActiveDocument();
    if (doc === null) {
      throw new Error('No active document');
    }
    var paragraphs = doc.getBody().getParagraphs();
  }
  
  for(var p in paragraphs) {
    if( ! paragraphs[p].getText())
      if( ! paragraphs[p].findElement(DocumentApp.ElementType.HORIZONTAL_RULE))
        if( ! paragraphs[p].findElement(DocumentApp.ElementType.PAGE_BREAK))
          if( ! paragraphs[p].findElement(DocumentApp.ElementType.INLINE_IMAGE))
            if( ! paragraphs[p].findElement(DocumentApp.ElementType.INLINE_DRAWING))
              if( ! paragraphs[p].isAtDocumentEnd())
                paragraphs[p].removeFromParent();
  }
}