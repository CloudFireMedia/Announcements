function rotateContent_() {

  //check to see if it's Sunday morning before Noon, if so, warn user
  
  var now = new Date();
  
  if (now.getDay() === 0 && now.getHours() < 12) {
  
    var title = 'Rotate Content';
    var prompt = Utilities.formatString(
      "\
This will remove the [Live Announcements Slides]\n\
(%s)\n\
that are queued for this Sunday's live announcements.\n\
\n\
Do you want to continue?\
",
      DriveApp.getFolderById(Config.get('SLIDES_FOLDER_DESTINATION_ID')).getUrl()
    );
    
    if (DocumentApp.getUi().alert(title, prompt, Browser.Buttons.YES_NO) !== 'YES') {
      return;
    }
  }
  
  moveThisToMaster();    // This (week 0) -> Master (check in)
  moveNextToThis();      // Next (week 1) -> This (week 0)
  moveDraftToNext();     // Draft (week 2) -> Next (week 1)
  moveMasterToDraft();   // Master -> Draft (check out)
  moveOldestToArchive(); // Master -> Archive (oldest)
  
  Comments_.update(config.lastTimeRotateRunText);

  return;
  
  // Private Functions
  // -----------------

  function moveThisToMaster(){
  
    var thisWeekId    = Config.get('ANNOUNCEMENTS_0WEEKS_SUNDAY_ID');
    var thisSundayDoc = DocumentApp.openById(thisWeekId);
    
    ///thisSundayHeading is dependent upon content - it may be better to grab the last placeholder paragraph
    var thisSundayHeading = getHeading(thisSundayDoc).trim();//gets the first non-empty paragraph text
    var thisSundayBody    = thisSundayDoc.getBody();
    
    var masterDocId = Config.get('ANNOUNCEMENTS_MASTER_SUNDAY_ID');
    var masterDoc   = DocumentApp.openById(masterDocId);
    var masterBody  = masterDoc.getBody();
    var startElem   = masterBody.findText(thisSundayHeading.replace('[','\\['));
    
    if ( ! startElem) {
      throw "Unable to find '" + thisSundayHeading + "' in Master doc"; 
    }
    
    var endElem = masterBody.findText(
      //condifer using: escapeGasRegExString(config.announcements.placeholder)
      config.announcements.placeholder
        .replace('{','\\{')
        .replace('[','\\[')
        .replace('|','\\|')
      , startElem
    );//might need some others
    
    var nextPageBreak   = masterBody.findElement(DocumentApp.ElementType.PAGE_BREAK, startElem);
    
    var startIndex      = startElem     && masterBody.getChildIndex(startElem.getElement().getParent());
    var endIndex        = endElem       && masterBody.getChildIndex(endElem.getElement().getParent());
    var pageBreakIndex  = nextPageBreak && masterBody.getChildIndex(nextPageBreak.getElement().getParent());
    var endIndexForReal = endIndex && pageBreakIndex && endIndex == pageBreakIndex ? endIndex : pageBreakIndex;
  
  //  throw endIndex +' : '+ pageBreakIndex +' : '+ endIndexForReal
    
    for (var i=endIndexForReal; i>=startIndex; i--) {
      masterBody.getChild(i).removeFromParent();
    }
    
    if(endIndexForReal != endIndex) {//oops, the pagebreak was contained in deleted content
      masterBody.insertPageBreak(startIndex);//so we add one back
    }
  
    //check-in This Sunday content
    for (var j=thisSundayBody.getNumChildren()-1; j>=0; --j) {//in reverse as each para is inserted before the last
    
      var thisSundayElem = thisSundayBody.getChild(j).copy();
      
      if (thisSundayElem.getType() == DocumentApp.ElementType.PARAGRAPH) {
        masterBody.insertParagraph(startIndex, thisSundayElem);//insert para from This Sunday
      } else if (thisSundayElem.getType() == DocumentApp.ElementType.TABLE) {
        masterBody.insertTable(startIndex, thisSundayElem);//untested
      } else if (thisSundayElem.getType() == DocumentApp.ElementType.INLINE_IMAGE) {
        //masterBody.insertImage(startIndex, thisSundayElem);//nope, ignore images, don't want them in master
      } else if (thisSundayElem.getType() == DocumentApp.ElementType.LIST_ITEM) {
        masterBody.insertListItem(startIndex, thisSundayElem);//untested
      } //else skip it
    }
  
    thisSundayBody.clear();//removes all content and formatting
    var sundayTitles = getWeekTitles();
    thisSundayDoc.setName(sundayTitles.thisSunday);
    
    format_master_(masterDoc);
    
  } // moveThisToMaster()
  
  function moveNextToThis(){
  
    var fromDoc = DocumentApp.openById(Config.get('ANNOUNCEMENTS_1WEEK_SUNDAY_ID'));
    var toDoc   = DocumentApp.openById(Config.get('ANNOUNCEMENTS_0WEEKS_SUNDAY_ID'));
    copyContent(fromDoc, toDoc);
      
    var sundayTitles = getWeekTitles();
  
    toDoc.setName(sundayTitles.thisSunday);
    fromDoc.setName(sundayTitles.nextSunday);
    fromDoc.getBody().clear();
  
    runAllFormattingFunctions_upcomingWeek_(toDoc);
  }
  
  function moveDraftToNext(){
  
    var sundayTitles = getWeekTitles();
    var title = getSundayOfMonthOrdinal(sundayTitles.dates.draftSunday) + ' Sunday of the month';
    var twoWeeksDocId = Config.get('ANNOUNCEMENTS_2WEEKS_SUNDAY_ID');
    var fromDoc = DocumentApp.openById(twoWeeksDocId);
    var toDoc   = DocumentApp.openById(Config.get('ANNOUNCEMENTS_1WEEK_SUNDAY_ID'));
    
    copyContent(fromDoc, toDoc);
  
    //add back the title back that is omitted from the draft version
    
    var body = toDoc.getBody();
    
    var titlePara = body.insertParagraph(0, sundayTitles.nextSunday)
      .setHeading(DocumentApp.ParagraphHeading.HEADING1);
      
    //the hrule is in a para by itself, not part of the title text  
    body.insertParagraph(1, '').appendHorizontalRule();
  
    toDoc.setName(sundayTitles.nextSunday);
    fromDoc.setName(sundayTitles.draftSunday);//just so there aren't two docs with the same name, even temporarily
    fromDoc.getBody().clear();
  
    runAllFormattingFunctions_oneWeekOut_(toDoc);
    
    return;
    
  } // moveDraftToNext()
  
  function moveMasterToDraft(){
  
    var masterDoc = DocumentApp.openById(Config.get('ANNOUNCEMENTS_MASTER_SUNDAY_ID'));
    var masterBody = masterDoc.getBody();
  
    var draftDoc = DocumentApp.openById(Config.get('ANNOUNCEMENTS_2WEEKS_SUNDAY_ID'));
    var draftBody = draftDoc.getBody();
  
    var sundayTitles = getWeekTitles();
    var draftTitle = sundayTitles.draftSunday;
    var nextDraftHeaderText = draftTitle.replace(/ +-? *Draft.*$/i, '');
  
    var startElem  = masterBody.findText(nextDraftHeaderText.replace('[','\\['));
    var startIndex = masterBody.getChildIndex(startElem.getElement().getParent());
    var hruleElem  = masterBody.findElement(DocumentApp.ElementType.HORIZONTAL_RULE, startElem);
    var hruleIndex = masterBody.getChildIndex(hruleElem.getElement().getParent());//normally the next element
    //sometimes the last para contains the pagebreak in which case it was being skipped
    //  var endElem    = masterBody.findElement(DocumentApp.ElementType.PAGE_BREAK, startElem);
    //this method ensures we get everything up to the next Sunday
    var endElem    = masterBody.findText(sundayTitles.nextSunday.replace('[','\\['));
  //  log(endElem.getElement().getParent().asText().getText());return;
    var endIndex   = masterBody.getChildIndex(endElem.getElement().getParent()) -1;
  
    draftBody.clear();
    draftDoc.setName(draftTitle);
  
    for (var i=endIndex; i>hruleIndex; i--){//skip the title block for the draft document
  //  for (var i=endIndex; i>=startIndex; i--){
      var element = masterBody.getChild(i).copy();//make a copy
      var type = element.getType();
  //    if(i > hruleIndex+1)//leave title and hrule, delete the rest
      masterBody.getChild(i).removeFromParent();
  //    masterBody.getChild(i).setAttributes({FOREGROUND_COLOR:'#0000ff'});///debug
      //copy to Draft doc
      if( type == DocumentApp.ElementType.PARAGRAPH ){
        var newPara = draftBody.insertParagraph(0, element);
        var pageBreak = newPara.findElement(DocumentApp.ElementType.PAGE_BREAK);
        if(pageBreak){//for then the pagebreak is in the final paragraph
          pageBreak.getElement().removeFromParent();
          masterBody.insertPageBreak(i);
        }
      }
      else if( type == DocumentApp.ElementType.TABLE )          draftBody.appendTable(element);
      else if( type == DocumentApp.ElementType.LIST_ITEM )      draftBody.appendListItem(element);
      else if( type == DocumentApp.ElementType.INLINE_IMAGE )   draftBody.appendImage(element);//untested
      else if( type == DocumentApp.ElementType.INLINE_DRAWING ) draftBody.appendImage(element);//untested
    }
    
    //add placeholder
    masterBody
    .insertParagraph(hruleIndex+1, config.announcements.placeholder)
    .setAlignment(DocumentApp.HorizontalAlignment.LEFT)
    .setAttributes({ITALIC : false, BOLD : false});
      
  //  format_master();///FIX
  }
  
  function moveOldestToArchive(){
  
    var archiveDocId = Config.get('ANNOUNCEMENTS_ARCHIVE_ID');
    var archiveDoc   = DocumentApp.openById(archiveDocId);  
    var archiveBody  = archiveDoc.getBody();
    
    var doc = DocumentApp.openById(Config.get('ANNOUNCEMENTS_MASTER_SUNDAY_ID'));
    var body = doc.getBody();
    var paragraphs = body.getParagraphs();
    
    //first, let's cleanup the doc to avoid moving blank paragraphs
    format_master_(doc);///FIX
    
    //get the index of the page to begin archive from - currently only the last page of announcements
    //could be changed to accept user date input or start from cursor location page
    for(var p=paragraphs.length-1; p>=0; p--) {//search from end for match - there's no inbuilt method for this
      if( ! paragraphs[p].getText().match(config.announcements.sundayPageRegEx)) continue;//we're looking for page titles, skip anything else
      var start = body.getChildIndex(paragraphs[p]);
      var currentChild = paragraphs[p];
      break;
    }
    
    if( ! start) {
      throw 'Unable to find any annoucement page';//this should never happen unless the regex is wrong
    }
    
    //start archiving
    var archiveInsertIndex = 0;
    
    while( ! currentChild.findText('(?i)\\[ *?END ?OF ?(?:DOC|DOCUMENT) *?]') ){//process all paragraphs up the [ END ]
    
      try {
        currentChild.removeFromParent();
      } catch(e) {
        /*errors if no END tag in doc, ignore*/
      };//could do this based on last child index but that keeps changing...yeah, leave it --Bob
      
      switch(currentChild.getType()){//different types are moved differently
          
        case DocumentApp.ElementType.PARAGRAPH:
          var what = currentChild.findElement(DocumentApp.ElementType.PAGE_BREAK) ? '':currentChild;//replace pagebreak with empty paragraph
          archiveBody
          .insertParagraph(archiveInsertIndex, what)
          .setAttributes(config.announcements.format.archive);
          break;
          
        case DocumentApp.ElementType.LIST_ITEM:
          //listitem glyphs retain their formatting...grrr
          //apparently there is no way at this time to change glyph formatting via scripting --Bob 2018-11-14
          archiveBody
          .insertListItem(archiveInsertIndex, currentChild)
          .setAttributes(config.announcements.format.archive);
          break;
          
        case DocumentApp.ElementType.TABLE:
          archiveBody
          .insertTable(archiveInsertIndex, currentChild)
          .setAttributes(config.announcements.format.archive);
          break;
          
        default:
          throw 'Unhandled child type "'+currentChild.getType()+'"';
      }
      
      archiveInsertIndex++;
      
      currentChild = body.getChild(start);//since the previous paragraph was removed, start is now the next paragraph
    }
  }
  
  function copyContent(fromDoc, toDoc){
  
    var fromBody = fromDoc.getBody();
    var toBody = toDoc.getBody();
    toBody.clear();//in case it isn't empty
    
    var totalElements = fromBody.getNumChildren();
    
    for (var j=0; j<totalElements; ++j ) {
    
      var element = fromBody.getChild(j).copy();
      var type = element.getType();
      
      if( type == DocumentApp.ElementType.PARAGRAPH )           toBody.appendParagraph(element);
      else if( type == DocumentApp.ElementType.TABLE )          toBody.appendTable(element);
      else if( type == DocumentApp.ElementType.LIST_ITEM )      toBody.appendListItem(element);
      else if( type == DocumentApp.ElementType.INLINE_IMAGE )   toBody.appendImage(element);
      else if( type == DocumentApp.ElementType.INLINE_DRAWING ) toBody.appendImage(element);
    }
    
    //since everything was appended, the intial empty paragraph is still there pushing the content down
    toBody.getChild(0).removeFromParent();
  
  } // rotateContent_.copyContent()  

} // rotateContent_()

//delete all content in the destination folder and copy service slides corresponding to the 
//events written in 'This Sunday's Announcements' doc to the destination folder

function copySlides_() {
  log('Start: copySlides_()');
  var doc = DocumentApp.openById(Config.get('ANNOUNCEMENTS_0WEEKS_SUNDAY_ID'));
  var text = doc.getBody().editAsText().getText();
  var announcementStartRegEx = /\[[^\|\]]+\|/g;
  var fileTitleArray = [];
  var fileTitleMatch;
  
  //removes the first and last characters, ie: the [ and | from the above (but oddly leaves the spaces)

  do {

    fileTitleMatch = announcementStartRegEx.exec(text);
    
    if (fileTitleMatch !== null) {
      //removes [, | and any spaces
      fileTitleArray.push(fileTitleMatch[0].replace(/(^. *| *.$)/g, ''));
    }
    
  } while (fileTitleMatch !== null) 
  
  var folderSource = DriveApp.getFolderById(Config.get('SLIDES_FOLDER_SOURCE_ID'));
  var folderDestination = DriveApp.getFolderById(Config.get('SLIDES_FOLDER_DESTINATION_ID'));
  if( ! folderSource) throw new Error('Unable to open source folder for slides.')
  if( ! folderDestination) throw new Error('Unable to open destination folder for slides.')
  if( ! (folderSource && folderDestination)) return;
  
  //delete existing files
  var filesDestination = folderDestination.getFiles();
  if( ! filesDestination) throw new Error('Unable to access slide destination folder')
  try{
    while (filesDestination.hasNext()) {
      var nextFile = filesDestination.next()
      nextFile.setTrashed(true);
      log('Deleted "' + nextFile.getName() + '" in "' + folderDestination.getId() + '"')
    }
  }catch(e){
    throw new Error('Unable to delete old files from slide destination folder')
  }
  
  log('fileTitleArray: ' + JSON.stringify(fileTitleArray))
  
  //copy this weeks files to folder
  var filesSource = folderSource.getFiles();
  try{
    while (filesSource.hasNext()) {
      var file = filesSource.next();
      var filename = file.getName();//could check by mimetype
      if (filename.match(/\.jpg|\.png/)){//only copy image files
        for (var f in fileTitleArray) {
          var filenameSansExt = filename.replace(/\.[^.]+$/i, '');//remove extension
          log('filenameSansExt: ' + filenameSansExt)
          var compareResult = compareStrings(filenameSansExt, fileTitleArray[f]);
          log('compareResult: ' + compareResult);
          if (compareResult > 0.75) {//check file name against announcement title
            file.makeCopy(folderDestination);
            log('Created copy of "' + filename + '" in "' + folderDestination.getId() + '"')
          }
        }
      }
    }
  } catch(e){
    throw new Error('Unable to copy slides to destination folder')
  }
  
  log('End: copySlides_()')
  
} // copySlides_()