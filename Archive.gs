function masterToArchive() {
  //firstly make sure we can access the archive doc
  if( ! config.files.announcements.archive) throw 'Archive Setup is not configured.  Archiving cannot run';
  try{ var archiveDoc = DocumentApp.openById(config.files.announcements.archive); }
  catch(e){throw 'Error accessing archive document. ('+config.files.announcements.archive+')';};
  var archiveBody = archiveDoc.getBody();

  var doc = DocumentApp.openById(config.files. announcements.archive);
  var body = doc.getBody();
  var paragraphs = body.getParagraphs();
  
  //first, let's cleanup the doc to avoid moving blank paragraphs
  format_master();
  
  //get the index of the page to begin archive from - currently only the last page of annoucements
  //could be changed to accept user date input or start from cursor location page
  for(var p=paragraphs.length-1; p>=0; p--) {//search from end for match - there's no inbuilt method for this
    if(paragraphs[p].getText().match(config.announcements.sundayPageRegEx)){//we're looking for page titles, skip anything else
      var start = body.getChildIndex(paragraphs[p]);
      var currentChild = paragraphs[p];
      break;
    }
  }
  if( ! start) throw 'Unable to find any annoucement page';//this should never happen unless the regex is wrong
  
  //start archiving
  var archiveInsertIndex = 0;
  while( ! currentChild.findText('(?i)\\[ *?END ?OF ?(?:DOC|DOCUMENT) *?]') ){//process all paragraphs up the [ END ]
    currentChild.removeFromParent();
    switch(currentChild.getType()){//different types are moved differently

      case DocumentApp.ElementType.PARAGRAPH:
        if(archiveInsertIndex > body.getNumChildren()) body.appendParagraph('');//sometimes the script fails with archiveInsertIndex being 1 too high, could seitch to an append but this seemed simpler. --Bob 2018-02-26
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

function formatFont_archive() {///change to use config formats
  var doc = DocumentApp.openById(config.files. announcements.archive);
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

