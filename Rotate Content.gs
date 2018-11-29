function rotateContent() {
  //check to see if it's Sunday morning before Noon, if so, warn user
  var now = new Date();
  if(now.getDay()==0 && now.getHours() < 12){
    var title = 'Rotate Content';
    var prompt = Utilities.formatString(
      "\
This will remove the [Live Announcements Slides]\n\
(%s)\n\
that are queued for this Sunday's live announcements.\n\
\n\
Do you want to continue?\
",
      DriveApp.getFolderById(config.files.slides.folderDestination).getUrl()
    );
    if(DocumentApp.getUi().alert(title, prompt, Browser.Buttons.YES_NO) != 'YES') return;
  }
  /*
  This   -> Master (check in)
  Next   -> This
  Draft  -> Next
  Master -> Draft (check out)
  Master -> Archive (oldest)
  */
  moveThisToMaster();
  moveNextToThis();
  moveDraftToNext();
  moveMasterToDraft();
  moveOldestToArchive();
}

function moveThisToMaster(){
  var thisSundayDoc     = DocumentApp.openById(config.files.announcements.upcoming);
  ///thisSundayHeading is dependent upon content - it may be better to grab the last placeholder paragraph
  var thisSundayHeading = getHeading(thisSundayDoc).trim();//gets the first non-empty paragraph text
  var thisSundayBody    = thisSundayDoc.getBody();
  
  var masterDoc  = DocumentApp.openById(config.files.announcements.master);
  var masterBody = masterDoc.getBody();
  var startElem  = masterBody.findText(thisSundayHeading.replace('[','\\['));
  if ( ! startElem) throw "Unable to find '" + thisSundayHeading + "' in Master doc"; //if the header wasn't found in the master doc, throw an error
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
  
  for (var i=endIndexForReal; i>=startIndex; i--){
//    masterBody.getChild(i).setAttributes({BACKGROUND_COLOR:'#ffff00'});
//    masterBody.getChild(i).setAttributes({FOREGROUND_COLOR:'#ff0000'});
    masterBody.getChild(i).removeFromParent();
  }
  if(endIndexForReal != endIndex)//oops, the pagebreak was contained in deleted content
    masterBody.insertPageBreak(startIndex);//so we add one back

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
  
  format_master();///FIX  
}

function moveNextToThis(){
  var fromDoc = DocumentApp.openById(config.files.announcements.oneWeek);
  var toDoc   = DocumentApp.openById(config.files.announcements.upcoming);
  copyContent(fromDoc, toDoc);
    
  var sundayTitles = getWeekTitles();

  toDoc.setName(sundayTitles.thisSunday);
  fromDoc.setName(sundayTitles.nextSunday);
  fromDoc.getBody().clear();

  runAllFormattingFunctions_upcomingWeek();
}

function moveDraftToNext(){
  var sundayTitles = getWeekTitles();
  var title = getSundayOfMonthOrdinal(sundayTitles.dates.draftSunday) + ' Sunday of the month';
  var fromDoc = DocumentApp.openById(config.files.announcements.twoWeeks);
  var toDoc   = DocumentApp.openById(config.files.announcements.oneWeek);

  copyContent(fromDoc, toDoc);

  //add back the title back that is omitted from the draft version
  var body = toDoc.getBody();
  var titlePara = body.insertParagraph(0, sundayTitles.nextSunday)
  .setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.insertParagraph(1, '')//the hrule is in a para by itself, not part of the title text
  .appendHorizontalRule();

  toDoc.setName(sundayTitles.nextSunday);
  fromDoc.setName(sundayTitles.draftSunday);//just so there aren't two docs with the same name, even temporarily
  fromDoc.getBody().clear();

  runAllFormattingFunctions_oneWeekOut(toDoc);
}

function moveMasterToDraft(){
  var masterDoc = DocumentApp.openById(config.files.announcements.master);
  var masterBody = masterDoc.getBody();

  var draftDoc = DocumentApp.openById(config.files.announcements.twoWeeks);
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
  //firstly make sure we can access the archive doc
//  if( ! config.files.announcements.archive) throw 'Archive Setup is not configured.  Archiving cannot run';
  try{ var archiveDoc = DocumentApp.openById(config.files. announcements.archive); }
  catch(e){throw 'Error accessing archive document. ('+config.files.announcements.archive+')';};
  var archiveBody = archiveDoc.getBody();
  
  var doc = DocumentApp.openById(config.files. announcements.master);
  var body = doc.getBody();
  var paragraphs = body.getParagraphs();
  
  //first, let's cleanup the doc to avoid moving blank paragraphs
  format_master();///FIX
  
  //get the index of the page to begin archive from - currently only the last page of annoucements
  //could be changed to accept user date input or start from cursor location page
  for(var p=paragraphs.length-1; p>=0; p--) {//search from end for match - there's no inbuilt method for this
    if( ! paragraphs[p].getText().match(config.announcements.sundayPageRegEx)) continue;//we're looking for page titles, skip anything else
    var start = body.getChildIndex(paragraphs[p]);
    var currentChild = paragraphs[p];
    break;
  }
  if( ! start) throw 'Unable to find any annoucement page';//this should never happen unless the regex is wrong
  
  //start archiving
  var archiveInsertIndex = 0;
  while( ! currentChild.findText('(?i)\\[ *?END ?OF ?(?:DOC|DOCUMENT) *?]') ){//process all paragraphs up the [ END ]
    try{currentChild.removeFromParent();} catch(e){/*errors if no END tag in doc, ignore*/};//could do this based on last child index but that keeps changing...yeah, leave it --Bob
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

//Redevelopment notes:  
  //                      - moveSlides function would be more accurately titled if it were called something like copySlides
  //                      - moveSlides should search the children of [Recurring Content](https://drive.google.com/drive/folders/1fy7tXK8ta6SpQAC2R1Xex35gIOW78SUp) folder  
  //                        as source folders in addition to the current source folder, which is [Pre- and Post- Service Slides](https://drive.google.com/drive/folders/0BzM8_MjdRURAXzlyVVNMdmpYLVU)
  //                      - moveSlides should be called on a timed trigger between 11:00:00 - 11:59:59 GMT-6 on Fridays (as midnight on Friday is the cutoff for staff to change content in This Sunday Announcements script)
  //                      - however, it can also be called manually, in which case... 
  //                        - the notification in lines 8 - 13 ONLY applies to the moveSlides function; no other functions should call this notification
  //                        - moveSlides should ONLY be able to run manually WITHOUT the notification in lines 8 - 13 ... 
  //                           IF the time now is NOT between between 00:00:01 GMT-6 Saturday - 12:00:00 GMT-6 Sunday;  
  //                           IF the filename of ['This Sunday' doc](https://docs.google.com/document/d/1U61THQS-Ktno-Ku1Jk6GX0UaSmsRgq4YGiJzePJleyo/edit) contains the date for this coming Sunday or today().
  //                           ELSE the notification should not appear

function moveSlides() {// delete all content in the destination folder and copy service slides corresponding to the events written in 'This Sunday's Announcements' doc to the destination folder
  log('Start: moveSlides()')
  var doc = DocumentApp.openById(config.files.announcements.upcoming);
  var text = doc.getBody().editAsText().getText();
  var announcementStartRegEx = /\[[^\|\]]+\|/g;
  var fileTitleArray = [];
  
  while (fileTitleMatch = announcementStartRegEx.exec(text))//removes the first and last characters, ie: the [ and | from the above (but oddly leaves the spaces)
  fileTitleArray.push( fileTitleMatch[0].replace(/(^. *| *.$)/g, '') );//removes [, | and any spaces
  
  var folderSource = DriveApp.getFolderById(config.files.slides.folderSource);
  var folderDestination = DriveApp.getFolderById(config.files.slides.folderDestination);
  if( ! folderSource) err('Unable to open source folder for slides.')
  if( ! folderDestination) err('Unable to open destination folder for slides.')
  if( ! (folderSource && folderDestination)) return;
  
  //delete existing files
  var filesDestination = folderDestination.getFiles();
  if( ! filesDestination) err('Unable to access slide destination folder')
  try{
    while (filesDestination.hasNext())
      filesDestination.next().setTrashed(true);
  }catch(e){
    err('Unable to delete old files from slide destination folder')
  }
  
  //copy this weeks files to folder
  var filesSource = folderSource.getFiles();
  try{
    while (filesSource.hasNext()) {
      var file = filesSource.next();
      var filename = file.getName();//could check by mimetype
      if (filename.match(/\.jpg|\.png/)){//only copy image files
        for (var f in fileTitleArray) {
          var filenameSansExt = filename.replace(/\.[^.]+$/i, '');//remove extension
          if (compareStrings(filenameSansExt, fileTitleArray[f]) > 0.75) {//check file name against announcement title
            file.makeCopy(folderDestination);
          }
        }
      }
    }
  }catch(e){err('Unable to copy slides to destination folder')}
  
  log('End: moveSlides()')
}

function copyContent(fromDoc, toDoc){
  var fromBody = fromDoc.getBody();
  var toBody = toDoc.getBody();
  toBody.clear();//in case it isn't empty
  
  var totalElements = fromBody.getNumChildren();
  for( var j=0; j<totalElements; ++j ) {
    var element = fromBody.getChild(j).copy();
    var type = element.getType();
    if( type == DocumentApp.ElementType.PARAGRAPH )           toBody.appendParagraph(element);
    else if( type == DocumentApp.ElementType.TABLE )          toBody.appendTable(element);
    else if( type == DocumentApp.ElementType.LIST_ITEM )      toBody.appendListItem(element);
    else if( type == DocumentApp.ElementType.INLINE_IMAGE )   toBody.appendImage(element);
    else if( type == DocumentApp.ElementType.INLINE_DRAWING ) toBody.appendImage(element);
  }
  
  toBody.getChild(0).removeFromParent();//since everything was appended, the intial empty paragraph is still there pushing the content down

}

function extractSundayTitle(doc){
  var search = doc.getBody().findText(config.announcements.sundayPagePattern);
  return search
  ? search.getElement().asText().getText()
  : null;
}

