function master_insertMonth() {///consider adding UI prompts

  //Note: This adds weeks for the next month. 
  //If the current month is missing Sundays, it adds them instead
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  
  //find last sunday
  var latestInsertedSundayRange = body.findText(config.announcements.sundayPagePattern);
  if(! latestInsertedSundayRange) throw 'Error finding previous Sunday.  Unable to continue.';
  var insertIndex = body.getChildIndex(latestInsertedSundayRange.getElement().getParent());
  var latestSundayShortcode = latestInsertedSundayRange.getElement().getParent().getText().match(/\d{2}\.\d{2}/)[0];//we know there's a match so get the data, no checks needed
  var latestInsertedSundayDate = getDateFromShortcode(latestSundayShortcode);
  
  //insertWeeksEndDate: finish the current month or add the next month
  var insertWeeksEndDate = latestInsertedSundayDate.getMonth() == dateAdd(latestInsertedSundayDate, 'week', 1).getMonth()
  ? getLastSundayOfMonth( latestInsertedSundayDate )                       //last Sunday of this month
  : getLastSundayOfMonth( dateAdd(latestInsertedSundayDate, 'month', 1) ); //last Sunday of next month
  
  var weeksToAdd = Math.round( ( insertWeeksEndDate.getTime() - latestInsertedSundayDate.getTime() ) / ( 7*24*60*60*1000 ) );//date diff / 1 week
  master_insertWeeks_(weeksToAdd);
}

function master_insertWeeks_(numWeeks){ //occasionally adds one week too many

  numWeeks = numWeeks || 1;
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();

  format_master_fixPageBreaksAndHeadings_(doc);//ensure everything is formatted correctly before we begin
  
  //find latest sunday
  var latestInsertedSundayRange = body.findText(config.announcements.sundayPagePattern);
  if(! latestInsertedSundayRange) throw 'Error finding previous Sunday.  Unable to continue.';
  var insertIndex = body.getChildIndex(latestInsertedSundayRange.getElement().getParent());
  var latestSundayShortcode = latestInsertedSundayRange.getElement().getParent().getText().match(/\d{2}\.\d{2}/)[0];//we know there's a match so get the data, no checks needed
  var latestInsertedSundayDate = getDateFromShortcode(latestSundayShortcode);

  for(var i=1; i<=numWeeks; i++){
  
    doc = DocumentApp.getActiveDocument();//have to do it this way or it errors with "too many changes"
    body = doc.getBody();
    
    var dateToInsert = new Date(new Date(latestInsertedSundayDate).setDate(latestInsertedSundayDate.getDate()+(7*i)));
    var newPageTitle = Utilities.formatDate(dateToInsert, Session.getScriptTimeZone(), "'[' MM.dd '] Sunday Announcements'");
    
    //add content in reverse at the insertIndex so we don't have to recalculate the index repeatedly
    //new page
    body.insertPageBreak(insertIndex);//inserts a pagebreak wrapped by a paragraph
    
    //add recurring content
    insertRecurringContent(dateToInsert, insertIndex, doc);
    
    //add page subtitle (ordinal Sunday)
    body
      .insertParagraph(insertIndex, getSundayOfMonthOrdinal(dateToInsert) + ' Sunday of the month')
      .setHeading(DocumentApp.ParagraphHeading.HEADING2);
      
    //add page title
    body.insertHorizontalRule(insertIndex);
    
    var pageTitle = body
      .insertParagraph(insertIndex, newPageTitle)
      .setHeading(DocumentApp.ParagraphHeading.HEADING1)//uses heading1 font settings, no need to add them here
      
    doc.setCursor(doc.newPosition(pageTitle, 0));//was for debug but I find it handy to end up at the top of the document
    doc.saveAndClose();//to avoid the "too many changes" error, close the doc then reopen it to add the next week
    
  } //next week
  
} // master_insertWeeks()

function insertRecurringContent(date, childIndex, doc) {

  date = date || new Date(); ///debug
  var body = doc.getBody();
  var recurringContent = getRecurringContent(doc);//{} with items parsed by directives i.e.: {first=[Paragraph],third=[Paragraph],date:{"2018-03-04":[Paragraph],}, ...}
  //  Logger.log(recurringContent);return;

  ////////////////////////////////////////////////////////////////////////////////////////////////////
  // Keys in recurringContent, present only when there are matching paragraphs, data is an array:
  // beforeFifth (fourth Sunday in a month with 5 Sundays), 
  // first, second, third, fourth, fifth, 
  // last (fourth or fifth Sunday), 
  // date (an object with date keys like '2018-04-22' which then contain an array of paragraphs), 
  // noMatchingDirective (in list but no known directive match)
  
  var dateOrdinal = getSundayOfMonthOrdinal(date).toLowerCase();//like: 'first'
  var isNextSundaySameMonth = date.getMonth() == dateAdd(date, 'week', 1).getMonth();//only false on the last Sunday of the month
  var isBeforeFifth = dateOrdinal=='fourth' && isNextSundaySameMonth;//only true if it's the fourth Sunday and there is a fifth Sunday
  var isLast = ! isNextSundaySameMonth;//only true on the last Sunday of the month
  
  //check in this order to avoid duplicates: specific date, beforeFifth, last, ordinal

  if(recurringContent.date){
    var pageDate = Utilities.formatDate(date, 0, 'yyyy-MM-dd');
    for(var d in recurringContent.date)//p like 2018-04-22
      if(pageDate == d)//if current page date matches recurringContent.date
        for(var p in recurringContent.date[d])
          insertParagraphCopy(body, recurringContent.date[d][p], childIndex);
  }

  if(recurringContent.beforeFifth && isBeforeFifth){
    for(var p in recurringContent.beforeFifth)
      insertParagraphCopy(body, recurringContent.beforeFifth[p], childIndex);
  }

  if(recurringContent.last && isLast){
    for(var p in recurringContent.last)
      insertParagraphCopy(body, recurringContent.last[p], childIndex);
  }

  if(recurringContent[dateOrdinal]){
    for(var p in recurringContent[dateOrdinal])
      insertParagraphCopy(body, recurringContent[dateOrdinal][p], childIndex);
  }
  
  //and we aren't doing anything with recurringContent.noMatchingDirective as we assume it's either a new directive or a typo

}

function insertParagraphCopy(body, paragraph, childIndex){
  var directivePattern = /<<.*?>>/;
  var para = body.insertParagraph(childIndex++, paragraph.copy());
  //para.replaceText(directivePattern, '');//this doesn't work?? no error, just no change
  para.setText(para.getText().replace(directivePattern, ' '));//replace directive with a single space
}

function getRecurringContent(doc){
  //get start and end indices for recurring content
  var body = doc.getBody();
  var recurringContent = [];
  var recurringContentRange = body.findText('(?i)\\[ RECURRING CONTENT ]');//this grabs the first one found, not there there should be more than one
  var recurringContentStart = body.getChildIndex(recurringContentRange.getElement().getParent());//the range is a text element, get its parent's index
  var recurringContentEnd   = body.getChildIndex(body.findElement(DocumentApp.ElementType.PAGE_BREAK, recurringContentRange).getElement().getParent())//get the index of the next pagebreak (hard breaks only)
  //Note: recurringContentEnd could use the first date page as the final index if there ever needs to be hard pages in the recurring content section
  
  //iterate over elements in found range collecting announcements, skipping blank paragraphs
  for(var i=recurringContentStart+1; i<=recurringContentEnd; i++){//+1 to skip the page title [ RECURRING CONTENT ]
    var text = body.getChild(i).asParagraph().getText();
    if(text)
      recurringContent.push(body.getChild(i));//if it's not blank, store the element
  }
  
  //arrange announcements by type in an object
  var output = {};
  for(var i in recurringContent) {
    //extract directives found between << angle braces >>
    var paragraph = recurringContent[i].asParagraph(); //var paragraph = body.findElement(elementType).getElement().asParagraph()
    var text = paragraph.getText();
    var directive = text.match(/<< *(.*?) *>>/);//returns the directive, trimmed, without the enclosures
    ///consider: var directives = directive.replace('before fifth', 'beforeFifth').split(' ') then use directives.indexOf('first') - naw: prefer regex simplicity
    if(!directive) continue;//a directive is required or we skip it
    directive = directive[1];//directive 0 = '<<directive>>' while directive[1] = 'directive' without the enclosures
    
    var beforeFifth = directive.match(/before *fifth/i);//check this first or it will get caught by nth Sunday test
    var beforeFifth = beforeFifth || directive.match(/when *a *fifth *Sunday *exists/i);//this is temporary to match old format
    directive = directive.replace(/before *fifth/i, '').replace(/when *a *fifth *Sunday *exists/i, '');//remove these so "fifth" doesn't trigger a false match later
    //<<announcement should be appended on the fourth Sunday of the month when a fifth Sunday exists in the same month>>
    //<<before fifth>>
    if(beforeFifth){
      if( ! output.beforeFifth) output.beforeFifth = [];
      output.beforeFifth.push(paragraph);
    }
    
    var ordinalSunday = directive.match(/(?:first|second|third|fourth|fifth|last)/gi);//note: this can have mutiple matches
    // <<announcement should be appended on the first Sunday of the month>> 
    // <<first>>, <<second>>, <<last>>, <<first Sunday>>, etc
    if(ordinalSunday){
      for(var o in ordinalSunday){//for each matching ordinal
        var oSunday = ordinalSunday[o].toLowerCase();
        if( ! output[oSunday]) output[oSunday] = [];//like: output['first'] = []
        output[oSunday].push(paragraph);
      }
    }
    
    var specificSunday = directive.match(/\d{2}\.\d{2}/g);//matches "03.10" - multiple possible
    //<<announcement should be appended on 03.04, 10.28>>
    //<<03.04, 10.28>> or [03.04] but not 3.04, 03-04
    if(specificSunday && specificSunday.length){
      for(var s in specificSunday){
        var dayt = Utilities.formatDate(getDateFromShortcode(specificSunday[s]), Session.getScriptTimeZone(), "yyyy-MM-dd");
        if( ! output.date) output.date = {};
        if( ! output.date[dayt]) output.date[dayt] = [];
        output.date[dayt].push(paragraph);
      }
    }
    
    //if nothing else caught it, then stick it here
    if( ! output.noMatchingDirective) output.noMatchingDirective = [];
    output.noMatchingDirective.push(paragraph);
    ///maybe notify someone... ////what think ye, chad? 
    
  }
  
//  Logger.log(output);
  return output;
}

function format_removeEmptyParagraphs() {
  format_removeEmptyParagraphs_();
}