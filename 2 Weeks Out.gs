///reorganize


//function transferText() {//on menu as Rotate Content
//  moveThisSunday_to_MasterSundayAnnouncements();
//  moveNextSunday_to_thisSunday();
//  moveSlides();
//}

//function moveThisSunday_to_MasterSundayAnnouncements() {
//  var thisSundayDoc = DocumentApp.openById(Config.get('ANNOUNCEMENTS_2WEEKS_SUNDAY_ID'));
//  var thisSundayHeading = getHeading(thisSundayDoc).trim(); //removes any trailing whitespace, to make the search more accurate
//  
//  //gets the "children" (paragraphs, tables, etc.) from the This Sunday doc
//  var thisSundayBody = thisSundayDoc.getBody();
//  var thisSundayChildrenCount = thisSundayBody.getNumChildren();
//  
//  //gets the "children" (paragraphs, tables, etc.) from the Master doc
//  var masterDoc = DocumentApp.openById(Config.get('ANNOUNCEMENTS_MASTER_SUNDAY_ID'));
//  var masterBody = masterDoc.getBody();
//  var masterChildren = masterBody.getNumChildren();
//  
//  //the code below searches through the Master children to find the heading.  When found, it inserts each of the This Sunday children in its place.
//  var found = false;
//  for (var i=0; i<masterChildren; ++i) {
//    var masterElem = masterBody.getChild(i);
//    var masterType = masterElem.getType();
//    var txt = masterElem.asText().getText();
//    if (txt.indexOf(thisSundayHeading) > -1) {//if the text matches the header, we've found it
//      masterElem.removeFromParent();//remove the header text from the master doc so there's not a duplicate when we insert the new stuff
//      found = true;
//      for (var j=thisSundayChildrenCount-1; j>=0; --j) {//copy This Sunday to Master, in reverse as each para is inserted before the last
//        var thisSundayElem = thisSundayBody.getChild(j).copy();
//        //var thisSundayType = thisSundayElem.getType();//unused
//        if (masterType == DocumentApp.ElementType.PARAGRAPH) {
//          //var masterElem1 = masterBody.getChild(i);///why?
//          //var masterType1 = masterElem1.getType();///why?
//          //var txt1 = masterElem1.asText().getText();
//          //txt1 = txt1.trim();
//          //if (txt1 == "~")///use a non-arcane placeholder
//          if(txt.trim() == config.announcements.placeholder) masterDoc.getChild(i).removeFromParent();//remove the placeholder
//          masterBody.insertParagraph(i, thisSundayElem);//insert para from This Sunday
//        } else if (masterType == DocumentApp.ElementType.TABLE) {
//          //masterBody.insertTable(i, thisSundayElem);//untested
//        } else if (masterType == DocumentApp.ElementType.INLINE_IMAGE) {
//          //masterBody.insertImage(i, thisSundayElem);//untested
//        } else if (masterType == DocumentApp.ElementType.LIST_ITEM) {
//          //masterBody.insertListItem(i, thisSundayElem);//untested
//        }
//      }
//      break;//stop looking once founs
//    }
//  }
//  if (found == false)
//    throw "Unable to find " + thisSundayHeading + " in Master doc"; //if the header wasn't found in the master doc, throw an error, and don't go on to clear out This Sunday/Next Sunday/etc.
//  
//}

//function moveNextSunday_to_thisSunday() {
//  var docNextSunday = DocumentApp.openById(Config.get('ANNOUNCEMENTS_1WEEK_SUNDAY_ID'));
//  var p = docNextSunday.getParagraphs();
//  
//  var docThisSunday = DocumentApp.openById(Config.get('ANNOUNCEMENTS_0WEEKS_SUNDAY_ID'));
//  docThisSunday.getBody().clear();
//  
//  var nextSundayTitle;
//  for (var i = 0; i < p.length; ++i) {
//    if (p[i].getText().trim().indexOf("[") == 0) {
//      nextSundayTitle = p[i].getText();
//      break;
//    }
//  }
//  if (nextSundayTitle != undefined) {
//    docThisSunday.setName(nextSundayTitle);
//  }
//  
//  var thisSundaybody = docThisSunday.getBody();
//  for (var i = 0; i < p.length; i++) {
//    thisSundaybody.appendParagraph(p[i].copy());
//  }
//  var thisSundayDocId = Config.get('ANNOUNCEMENTS_0WEEKS_SUNDAY_ID');
//  applyFormattingToHeading(thisSundayDocId);
//  applyFormattingBody(thisSundayDocId);
//  //Draft To Next Sunday
//  DraftToNextSunday();
//}

//function DraftToNextSunday() {
//  var draftDoc = DocumentApp.openById(Config.get('ANNOUNCEMENTS_0WEEKS_SUNDAY_ID'));
//  var documentName = draftDoc.getName();
//  var arr = documentName.split("-");
//  var paraFirst = arr[0];
//  
//  var docNextSunday = DocumentApp.openById(Config.get('ANNOUNCEMENTS_1WEEK_SUNDAY_ID'));
//  docNextSunday.setName(paraFirst);
//  docNextSunday.getBody().clear();
//  
//  var nextSundaybody = docNextSunday.getBody();
//  nextSundaybody.appendParagraph(paraFirst);
//  nextSundaybody.appendHorizontalRule();
//  var p = draftDoc.getBody().getParagraphs();
//  for (var i = 0; i < p.length; i++) {
//    var foundImage = false;
//    for (var j = 0; j < p[i].getNumChildren(); ++j) {
//      var child = p[i].getChild(j);
//      if (child.getType() == DocumentApp.ElementType.INLINE_IMAGE) {
//        foundImage = true;
//        break;
//      }
//    }
//    if (foundImage) continue;
//    nextSundaybody.appendParagraph(p[i].copy());
//  }
//  applyFormattingToHeading(Config.get('ANNOUNCEMENTS_1WEEK_SUNDAY_ID'));
//  masterToDraft();
//}

function applyFormattingBody(id) {
  var style = {};
  style[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] =
    DocumentApp.HorizontalAlignment.LEFT;
  style[DocumentApp.Attribute.FONT_FAMILY] = 'Lato';
  style[DocumentApp.Attribute.FONT_SIZE] = 9;
  style[DocumentApp.Attribute.FOREGROUND_COLOR] = "#aeaeb2";
  var doc = DocumentApp.openById(id).getBody();
  doc.setAttributes(style);
}

function applyFormattingToHeading(id) {
  var style = {};
  style[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] =
    DocumentApp.HorizontalAlignment.CENTER;
  style[DocumentApp.Attribute.FONT_FAMILY] = 'Lato';
  style[DocumentApp.Attribute.FONT_SIZE] = 9;
  style[DocumentApp.Attribute.BOLD] = true;
  style[DocumentApp.Attribute.FOREGROUND_COLOR] = "#aeaeb2";
  var doc = DocumentApp.openById(id);
  var firstPara = getFirstParagraph(doc);
  firstPara.setAttributes(style);
}

function getHeading(doc) {
  var p = doc.getBody().getParagraphs();
  for (i=0; i<p.length; i++)
    if (p[i].getText() != "")
      return p[i].getText();
  return 'Not Found';
}

function getFirstParagraph(doc) {
  var p = doc.getBody().getParagraphs();
  for (i=0; i<p.length; i++)
    if (p[i].getText() != "")
      return p[i];
  return 'Not Found';
}

//function getNextDraft_OLD() {
//  //this gets the name of the current draft and adds 7 days
//  var docSunday = DocumentApp.openById(Config.get('ANNOUNCEMENTS_2WEEKS_SUNDAY_ID'));
//  var documentName = docSunday.getName();
//log(documentName)
//  var arrnew = documentName.split("-");
//  var paraFirst = arrnew[0];
//  var str = arrnew[0];
//  //return;
//  var arr = str.split("[");
//  var arr1 = arr[1].split("]");
//  var datestr = arr1[0].trim();
//  var arr2 = datestr.split(".");
//  var currentDate = "2017/" + arr2[0] + "/" + arr2[1];
//  datestr.replace(".", ":");
//  var today = new Date(currentDate); //Today's Date
//  var next = today.setDate(today.getDate() + 7);
//  var dateFormated = today.toISOString().substr(0, 10);
//  var arr3 = dateFormated.split("-");
//  var nextdraftDate = arr3[1] + "." + arr3[2];
//  nextdraftDate = '[ ' + nextdraftDate + ' ]' + arr1[1];
//  if (arr[1] != "") {
//    var nextdraftDate1 = nextdraftDate;
//    nextdraftDate = nextdraftDate + "-" + arrnew[1];
//  }
////  docSunday.setName(nextdraftDate);
//log(nextdraftDate1);
//  return nextdraftDate1;
//}

//function getNextDraft() {
//  
//  return fDate(
//    dateAdd(getUpcomingSunday(), 'week', 2), 
//    "[ MM.dd ] 'Sunday Announcements - Draft Document'"
//  );
//
//  //forget all this jargon
//  
//  //this gets the name of the current draft and adds 7 days
//  var arrnew = documentName.split("-");
//  var paraFirst = arrnew[0];
//  var str = arrnew[0];
//  //return;
//  var arr = str.split("[");
//  var arr1 = arr[1].split("]");
//  var datestr = arr1[0].trim();
//  var arr2 = datestr.split(".");
//  var currentDate = "2017/" + arr2[0] + "/" + arr2[1];
//  datestr.replace(".", ":");
//  var today = new Date(currentDate); //Today's Date
//  var next = today.setDate(today.getDate() + 7);
//  var dateFormated = today.toISOString().substr(0, 10);
//  var arr3 = dateFormated.split("-");
//  var nextdraftDate = arr3[1] + "." + arr3[2];
//  nextdraftDate = '[ ' + nextdraftDate + ' ]' + arr1[1];
//  if (arr[1] != "") {
//    var nextdraftDate1 = nextdraftDate;
//    nextdraftDate = nextdraftDate + "-" + arrnew[1];
//  }
////  docSunday.setName(nextdraftDate);
//log(nextdraftDate1);
//  return nextdraftDate1;
//}

//var stoopidGlobalVarThatNeedsReplacedASAP = ""; /////////////////////////////////////////////////make it go away!
//function masterToDraft() {
//  
//  var docDraft = DocumentApp.openById(Config.get('ANNOUNCEMENTS_0WEEKS_SUNDAY_ID'));
//  docDraft.getBody().clear();
//  var draftBody = docDraft.getBody();
//  var nextDraftHeaderText = getNextDraft().trim();
//  var docMaster = DocumentApp.openById(Config.get('ANNOUNCEMENTS_MASTER_SUNDAY_ID'));
//  var masterBody = docMaster.getBody();
//  var numChildren = masterBody.getNumChildren();
//  
//  var foundDate = false;
//  var insertedTilde = false;
//  var elementsToRemove = [];
//  for (var e = 0; e < numChildren; ++e) {
//    var child = masterBody.getChild(e);
//    var childType = child.getType();
//    if (childType == DocumentApp.ElementType.PARAGRAPH) {
//      var text = child.asParagraph().copy().editAsText().getText();
//      if (text.indexOf(nextDraftHeaderText) != -1) {
//        stoopidGlobalVarThatNeedsReplacedASAP = e;
//        foundDate = true; //found a match for next week's date
//        continue;
//      } else if ((text.indexOf("[") != -1) && (text.indexOf("]") != -1) && (text.indexOf("Sunday Announcements") != -1) && foundDate) {
//        break; //we've reached the next section; stop searching
//      }
//    }
//    if (foundDate && childType == DocumentApp.ElementType.PARAGRAPH) { //append all paragraphs that don't contain a horizontal line
//      var paragraph = child.asParagraph();
//      var foundHorizontalLine = false;
//      for (var pe = 0; pe < paragraph.getNumChildren(); ++pe) {
//        if (paragraph.getChild(pe).getType() == DocumentApp.ElementType.HORIZONTAL_RULE) {
//          foundHorizontalLine = true;
//          break;
//        }
//      }
//      if (!foundHorizontalLine) {
//        draftBody.appendParagraph(paragraph.copy());
//        if (!insertedTilde) { //replace everything we've copied from master with a tilde
//          paragraph.setAlignment(DocumentApp.HorizontalAlignment.LEFT).setText(config.announcements.placeholder);
//          insertedTilde = true;
//        } else {
//          elementsToRemove.push(child);
//        }
//      }
//    }
//  }
//  elementsToRemove = elementsToRemove.reverse();
//  for (e in elementsToRemove) {
//    masterBody.removeChild(elementsToRemove[e]);
//  }
//  
//  stoopidGlobalVarThatNeedsReplacedASAP = stoopidGlobalVarThatNeedsReplacedASAP + 2;
//  masterBody.insertPageBreak(stoopidGlobalVarThatNeedsReplacedASAP);
//}

function draft_callFunctions() {
  draft_removeRowReferences();
  draft_mergeParagraph();
  //draft_removeSpaces()
  //draft_removeBlankPara();
  draft_formatParagraph();
  draft_boldBetweenSquareBrackets();
  draft_highlightStaffSponsorNames();
}

function draft_removeRowReferences() {
  var doc = DocumentApp.openById(Config.get('ANNOUNCEMENTS_2WEEKS_SUNDAY_ID'));
  var body = doc.getBody();
  var paras = body.getParagraphs();
  var length = paras.length;
  for (i = 0; i < length; ++i) {
    var str = paras[i].getText();
    if (str.indexOf("of the month") == -1) {
      var arr = str.split("]");
      var str1 = arr[0];
      const regex = /\|.*\|/g;
      const subst = '|';
      
      str1 = str1.replace(regex, subst);
      var found = 0;
      
      if (arr[1] != "" && arr[1] != undefined) {
        found = 1;
        var result = str1 + "]" + arr[1];
      } else {
        if (str1 != undefined) {
          found = 1;
          if (str.indexOf("]") != -1) {
            var result = str1 + "]";
          } else {
            var result = str1;
          }
        }
      }
      
      if (arr[2] != "" && arr[2] != undefined) {
        found = 1;
        var result = result + "]" + arr[2];
      }
      
      if (found == 1 && result != "") {
        paras[i].setText(result);
      }
    }
  }
}

function draft_removeSpaces() {
  var doc = DocumentApp.openById(Config.get('ANNOUNCEMENTS_2WEEKS_SUNDAY_ID'));
  var body = doc.getBody();
  var paras = body.getParagraphs();
  var length = paras.length;
  var cnt2 = 0;
  //Merge Paragraph
  for (i = 0; i < length; ++i) {
    var paraText = paras[i].getText();
    if (paraText != "") {
      if (cnt2 > 1) {
        paraText = paraText.replace(/\s\s+/g, ' ');
        
        paraText = paraText.replace(/[\r\n]/g, "\n");
        paraText = paraText.replace(". >>", ".\n>>");
        paraText = paraText.replace("! >>", "!\n>>");
        paras[i].setText(paraText);
      }
    }
    cnt2++;
  }
}

function draft_removeBlankPara() {
  var doc = DocumentApp.openById(Config.get('ANNOUNCEMENTS_2WEEKS_SUNDAY_ID'));
  var body = doc.getBody();
  var paras = body.getParagraphs();
  var length = paras.length;
  var cnt2 = 0;
  //Merge Paragraph
  for (i = 0; i < length; ++i) {
    var paraText = paras[i].getText();
    if (paraText == "") {
      if (cnt2 > 1) {
        try {
          paras[i].removeFromParent();
        } catch (e) {}
      }
      cnt2++;
    }
  }
}

function draft_mergeParagraph() {
  var doc = DocumentApp.openById(Config.get('ANNOUNCEMENTS_2WEEKS_SUNDAY_ID'));
  var body = doc.getBody();
  var paras = body.getParagraphs();
  var textLocation = {};
  var i;
  var length = paras.length;
  var cnt = 0;
  //Merge Paragraph
  for (i = 0; i < length; ++i) {
    var paraText = paras[i].getText();
    if (paraText != "") {
      if (cnt != 0) {
        paraText = paraText.trim();
        if (paraText.indexOf(">>") == 0 || paraText.indexOf(">>") == 1 || paraText.indexOf(">>") == 2 || paraText.indexOf(">>") == 3 || paraText.indexOf(">>") == 4) {
          var paraTextPrevious = paras[i - 1].getText();
          paraTextPrevious = paraTextPrevious.replace(/[\r\n]/g, "");
          paraText = paraText.replace(/[\r\n]/g, "");
          //paraTextPrevious=paraTextPrevious.replace("\n","");
          var newParaText = paraTextPrevious + "\n" + paraText;
          paras[i - 1].setText(newParaText);
          if (length - 1 == i) {
            paras[i].setText(".");
          } else {
            try {
              paras[i].removeFromParent();
            } catch (e) {}
          }
        }
      }
      cnt++;
    }
  }
}

function draft_formatParagraph() {
  var doc = DocumentApp.openById(Config.get('ANNOUNCEMENTS_2WEEKS_SUNDAY_ID'));
  var body = doc.getBody();
  var paras = body.getParagraphs();
  //Format Heading
  var cnt = 0;
  for (i = 0; i < paras.length; ++i) {
    var paraText = paras[i].getText();
    if (paraText != "") {
      if (cnt == 0) {
        var headingStyle = {};
        headingStyle[DocumentApp.Attribute.BOLD] = false;
        headingStyle[DocumentApp.Attribute.FONT_SIZE] = 9;
        headingStyle[DocumentApp.Attribute.ITALIC] = true;
        headingStyle[DocumentApp.Attribute.FONT_FAMILY] = 'Lato';
        paras[i].setAttributes(headingStyle);
        paras[i].setHeading(DocumentApp.ParagraphHeading.HEADING2);
      } else {
        var bodyStyle = {};
        bodyStyle[DocumentApp.Attribute.BOLD] = false;
        bodyStyle[DocumentApp.Attribute.FONT_SIZE] = 9;
        bodyStyle[DocumentApp.Attribute.FONT_FAMILY] = 'Lato';
        bodyStyle[DocumentApp.Attribute.UNDERLINE] = false;
        bodyStyle[DocumentApp.Attribute.LINE_SPACING] = 1.5;
        bodyStyle[DocumentApp.Attribute.SPACING_BEFORE] = 12;
        bodyStyle[DocumentApp.Attribute.SPACING_AFTER] = 0;
        bodyStyle[DocumentApp.Attribute.FOREGROUND_COLOR] = '#50505A';
        bodyStyle[DocumentApp.Attribute.BACKGROUND_COLOR] = '#FFFFFF';
        paras[i].setAttributes(bodyStyle);
        paras[i].setAlignment(DocumentApp.HorizontalAlignment.LEFT);
      }
      cnt++;
    } else {
      if (cnt > 1) {
        try {
          paras[i].removeFromParent();
        } catch (e) {}
        
      }
    }
  }
}

function draft_boldBetweenSquareBrackets() {
  var startTag = "[\[]";
  var endTag = "[\]]";
  var doc = DocumentApp.openById(Config.get('ANNOUNCEMENTS_2WEEKS_SUNDAY_ID'));
  var body = doc.getBody();
  var para = body.getParagraphs();
  var i = 0;
  for (i in para) {
    var from = para[i].findText(startTag);
    var to = para[i].findText(endTag, from);
    if ((to != null && from != null) && ((to.getStartOffset()) - (from.getStartOffset() + startTag.length) > 0)) {
      para[i].editAsText().setBold(from.getStartOffset(), to.getStartOffset(), true);
    }
  }
}

function draft_highlightStaffSponsorNames() {

  var doc = DocumentApp.openById(Config.get('ANNOUNCEMENTS_2WEEKS_SUNDAY_ID'));
  var body = doc.getBody();
  var textToHighlight = checkStaffDataSheet();
  var highlightStyle = {};
  highlightStyle[DocumentApp.Attribute.BACKGROUND_COLOR] = '#FCFC00';
  highlightStyle[DocumentApp.Attribute.BOLD] = true;
  var paras = body.getParagraphs();
  var textLocation = {};
  var i;
  
  for (i = 0; i < paras.length; ++i) {
    for (d in textToHighlight) {
      var textToFind = textToHighlight[d];
      textLocation = paras[i].findText(textToFind);
      if (textLocation != null && textLocation.getStartOffset() != -1) {
        textLocation.getElement().setAttributes(textLocation.getStartOffset(), textLocation.getEndOffsetInclusive(), highlightStyle);
      }
    }
  }
  
  return;
  
  // Private Functions
  // -----------------
  
  function checkStaffDataSheet() {
    var staffArray = [];
    var staffSheet = SpreadsheetApp.openById(Config.get('STAFF_DATA_GSHEET_ID'));
    var lastRow = staffSheet.getLastRow();
    var staffRows = staffSheet.getSheetValues(3, 1, lastRow, 2);
    var staffEmails = staffSheet.getSheetValues(3, 9, lastRow, 1);
    var staffFromSheet = "";
    
    var cnt = 0;
    for (var row in staffRows) {
      if (staffRows[row][0] != "") {
        staffFromSheet = staffRows[row][0] + ' ' + staffRows[row][1];
        staffArray.push(staffFromSheet);
        
      }
      cnt++;
    }
    return staffArray;
  }
}

function sendMailFunction_() {

  var docSunday = DocumentApp.getActiveDocument();
  
  if (docSunday === null) {
    docSunday = DocumentApp.openById(TEST_GDOC_ID);
  }
  
  var documentName = docSunday.getName();
  var arrnew = documentName.split("-");
  var paraFirst = arrnew[0];
  var str = arrnew[0];
  var arr = str.split("[");
  var arr1 = arr[1].split("]");
  var datestr = arr1[0].trim();
  checkName(datestr);
  
  return
  
  // Private Functions
  // -----------------

  function checkName(dateToSend) {
  
    var emailList = "";
    var staffToEmail = makestaffMailList();
    
    var comments = ('' + getCommentsFromDocument()).toUpperCase();
    var emailListArray = [];
    
    for (data in staffToEmail) {
    
      var emailAddress = staffToEmail[data];
      var staffName = emailAddress[0];
      var staffEmail = emailAddress[1];
      
      if (staffEmail != "") {
        var res = findNameInDraft(staffName);
        if (res || comments.indexOf(('' + staffName).toUpperCase()) > -1) //if staffname is in text || staffname is in a comment on the doc
        emailListArray.push(staffEmail);
      }      
    }
    
    var commentArryRes = driveAPI();    
    var array3 = arrayUnique(emailListArray.concat(commentArryRes));
    var uniqueArray = (uniq(array3));
    for (em in uniqueArray) {
      emailList += uniqueArray[em] + ",";
    }
    
    sendDraftMailFinal(emailList, dateToSend);
    
    return;
    
    // Private Functions
    // -----------------

    function getCommentsFromDocument() {
      var document_id = Config.get('ANNOUNCEMENTS_2WEEKS_SUNDAY_ID');
      var comments_list = Drive.Comments.list(document_id);
      var comments = "";
      
      for (var i = 0; i < comments_list.items.length; i++) {
        if (comments_list.items[i].status == "open" && comments_list.items[i].deleted == false) {
          comments += "  " + (comments_list.items[i].content)
        }
      }
      
      return comments;
    }

    function driveAPI() { ///rename this to something that makes sense with what it's doing.
    
      var emailList = "";
      var comments_list = Drive.Comments.list(Config.get('ANNOUNCEMENTS_2WEEKS_SUNDAY_ID')).items;
      var commentArray = [];

      if (comments_list.length > 0) {
      
        var newstr = JSON.stringify(comments_list);
        var arr1 = newstr.split('{"content":"');
        var arr2 = arr1[1].split('","htmlContent');
        var att = arr2[0].split(",");
        
        var len = att.length;
        
        if (len == 1) {
          var att = arr2[0].split(";");
        }
                
        for (n in att) {
          var nameToSearch = att[n];
          var sheet = SpreadsheetApp.openById(Config.get('STAFF_DATA_GSHEET_ID')).getActiveSheet();
          var dataRange = sheet.getDataRange();
          var values = dataRange.getValues();
          for (var i = 2; i < values.length; i++) { //skip header row
            var name = values[i][0] + " " + values[i][1];
            nameToSearch = nameToSearch.trim();
            name = name.trim();
            if (nameToSearch == name) {
              commentArray.push(values[i][8]);
            }
          }
        }
      }
      
      return commentArray;
      
    } // sendMailFunction_.checkName.driveAPI()    
    
  } // sendMailFunction_.checkName()
  
} // sendMailFunction_()

function sendDraftMailFinal(emailList, dateToSend) {

  if (emailList === '') {
    return;
  }

  var sundayAnnouncementsDraftDocumentUrl = DocumentApp.openById(Config.get('ANNOUNCEMENTS_2WEEKS_SUNDAY_ID')).getUrl();
  var subject = Utilities.formatString("Please Review: [ %s ] Sunday Announcements draft", dateToSend);
  var body = Utilities.formatString("Dear Event Sponsor: <br><br> \
Please review the document linked below regarding promotion for your upcoming event.<br><br>\
You are invited to suggest changes to your event's promotion by typing directly in the document by Friday EOD. <br><br>\
Thank you!<br><br>--<br>\
<a href='%s'>[ %s ] Sunday Announcement draft</a>", 
                                    sundayAnnouncementsDraftDocumentUrl,
                                    dateToSend
                                   );
  emailList = emailList.replace(/\,$/, ''); //should build the list without the trailing comma in the first place
  
  MailApp.sendEmail({
    to: emailList,
    subject: subject,
    htmlBody: body
  });
}

function makestaffMailList() {///redo this to use getDataRange() then reduce to get needed output
  var staffSheet = SpreadsheetApp.openById(Config.get('STAFF_DATA_GSHEET_ID'));
  var numRows = staffSheet.getLastRow();
  var staffRows = staffSheet.getSheetValues(3, 1, numRows, 2);
  var staffEmails = staffSheet.getSheetValues(3, 9, numRows, 1);
  var staffFromSheet = [];
  
  for (var row in staffRows) {
    staffFromSheet.push([staffRows[row][0] + ' ' + staffRows[row][1], staffEmails[row][0]]);
  }
  return staffFromSheet;
}

function findNameInDraft(staffName) {
  var check = 0;
  staffName = staffName.trim();
  if (staffName != "") {
    var doc = DocumentApp.openById(Config.get('ANNOUNCEMENTS_2WEEKS_SUNDAY_ID'));
    var body = doc.getBody();
    var textToHighlight = staffName;
    var paras = body.getParagraphs();
    var textLocation = {};
    var i;
    var result = "";
    for (i = 0; i < paras.length; ++i) {
      textLocation = paras[i].findText(textToHighlight);
      if (textLocation != null && textLocation.getStartOffset() != -1) return true;
    }
  }
  return false;
}


function reorderParagraphs() {
  var doc = DocumentApp.openById(Config.get('ANNOUNCEMENTS_2WEEKS_SUNDAY_ID'));
  var pa = doc.getBody().getParagraphs();
  var npa = [];
  var npi = 0;
  var md = '';
  var mdr = /^(\[[^\]\|]+\|[^\]]+\][^01-9]+)([01-9]+\.[01-9]+)([^01-9a-z]+)/gi;
  var mdr2 = '^\\[[^\]\\|]+\\|[^\\]]+\\][^01-9]+([01-9]+\.[01-9]+)[^01-9a-zA-Z]+';
  for (var pi = 0; pi < pa.length; pi++) {
    var pt = pa[pi];
    var ptt = pt.getText();
    npa[npi] = {};
    npa[npi].paragraph = pt;
    npa[npi].sort = false;
    if (ptt.match(mdr)) {
      md = mdr.exec(ptt);
      npa[npi].md = parseFloat(md[2]);
      //npa[npi].paragraph.replaceText(mdr2,md[1]);
      //ptt=npa[npi].paragraph.getText();
      npa[npi].text = ptt;
      npa[npi].sort = true;
      //npa[npi].md='';//parseFloat(md[0].replace(/^\[[^\]\|]+\|[^\]]+\][^01-9]+([01-9]+\.[01-9]+)[^01-9]/gi,'$1'));
    }
    npi++;
  }
  var sdfa = '112';
  var tp = '';
  for (npi = 0; npi < npa.length; npi++) {
    if (npa[npi].sort == true) {
      var tpi = npi;
      for (var npi2 = npi + 1; npi2 < npa.length; npi2++)
        if (npa[npi2].sort == true) {
          if (npa[npi2].md < npa[tpi].md) {
            tpi = npi2;
          }
        }
      if (tpi != npi) {
        tp = npa[npi];
        npa[npi] = npa[tpi];
        npa[tpi] = tp;
      }
    }
  }
  doc.getBody().appendParagraph("");
  var asdf = '';
  for (var pi = 0; pi < pa.length; pi++) {
    //pa[pi].removeFromParent();//doc.getBody().removeChild(pa[pi]);
    doc.getBody().removeChild(pa[pi]);
  }
  var asdff = '';
  for (npi = 0; npi < npa.length; npi++) {
    doc.getBody().appendParagraph(npa[npi].paragraph);
  }
  doc.getBody().removeChild(doc.getBody().getParagraphs()[0]);
  var asdffff = '';
}



/* Redevelopment note: it would be great if this script would also reference 
'This Sunday's Announcements' as well as 'Next Sunday's Announcements' as the
source doc. If two matches are found, the more recently edited version should be used 
(i.e.the version on 'Next Sunday's Announcements'). Even better would be if this script
could ALSO reference all pages dated 8 weeks prior to this Sunday on the Master Doc 
for the same reason, with the same rule about more recently edited content taking priority.
*/
//This function will use Regex to match events
function matchEvents() {
  
  var doc = DocumentApp.openById(Config.get('ANNOUNCEMENTS_1WEEK_SUNDAY_ID'));
  var opa = doc.getBody().getParagraphs(); //opa contains event text
  
  var draftdoc = DocumentApp.openById(Config.get('ANNOUNCEMENTS_2WEEKS_SUNDAY_ID'));
  //will return Paragraph array
  var pa = draftdoc.getBody().getParagraphs(); //contains event and style content
  var npa = [];
  var npi = 0;
  var md = '';
  //var mdr=/^\[[^\]\|]+\|[^\]]+\][^01-9]+([01-9]+\.[01-9]+)[^01-9]/gi ;
  //var mdr2=/^\[[^\]\|]+\|[^\]]+\][^01-9]+([01-9]+\.[01-9]+)[^01-9].*[a-z]+/gi ;
  
  //Build regex object
  var mdr = /^\[([^\]\|]+\|[^\]]+)\]/gi;
  var mdr2 = /^\[[^\]\|]+\|[^\]]+\].*[a-z]+/gi;
  var mdrblank = /[a-z]+/gi;
  
  //DRAFT DOC:  iterate through the paragraph array
  draftdoc.getBody().appendParagraph(" ");
  for (var pi = 0; pi < pa.length; pi++) {
    var pt = pa[pi];
    var ptt = pt.getText(); //get paragraph text as String
    npa[npi] = {};
    npa[npi].paragraph = pt;
    npa[npi].sort = false;
    
    //perform the regex match
    if ((ptt.match(mdr)) && !(ptt.match(mdr2))) {
      
      npa[npi].text = ptt; //store event text
      npa[npi].replace = true;
      md = mdr.exec(ptt); //store matched text in md
      npa[npi].event = md[0]; //and in event 
      
      //work on EVENTS document
      for (var opi = 0; opi < opa.length; opi++) {
        //if event matches
        if (opa[opi].getText().indexOf(npa[npi].event) >= 0) {
        
          //overwrite the event paragraph fetched from DRAFT doc with the one from EVENTS doc
          //npa[npi].paragraph=opa[opi].copy();
          npa[npi].paragraph = pt.copy();
          
          //text from reference document
          var txt = opa[opi].getText();
          
          //in the reference paragraph , get only text succeeding the first semi colon (;)
          // txt = txt.substring(txt.indexOf(';') + 1, txt.length);
          txt = txt.replace(npa[npi].event, '');
          //get current 
          //--------------------------------
          var body = draftdoc.getBody();
          var txtToReplace = ptt.substring(0, ptt.indexOf(';') + 1);
          //var ptText = txtToReplace + txt;
          var ptText = ptt + txt;
          
          var limit = ptt.length - 1;
          var att = pt.getAttributes();
          var par = body.insertParagraph(pi, ptText);
          
          var style = {};
          style[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = pt.getAlignment();
          style[DocumentApp.Attribute.FONT_FAMILY] = 'Lato';
          style[DocumentApp.Attribute.FONT_SIZE] = 9;
          style[DocumentApp.Attribute.BOLD] = false;
          style[DocumentApp.Attribute.LINE_SPACING] = 1.5;
          
          style[DocumentApp.Attribute.SPACING_AFTER] = 10;
          style[DocumentApp.Attribute.SPACING_BEFORE] = 10;
          style[DocumentApp.Attribute.FOREGROUND_COLOR] = '#58585A';
          par.setAttributes(style);
          par.editAsText().setBold(0, limit, true); //new
          pt.removeFromParent();
          //-------------------------------          
        }
      }
    }
    npi++;
  }
}

function removeShortStartDates_OLD() {
  var doc = DocumentApp.openById(Config.get('ANNOUNCEMENTS_2WEEKS_SUNDAY_ID'));
  var pa = doc.getBody().getParagraphs();
  var mdr = /^(\[[^\]\|]+\|[^\]]+\][^01-9]+)([01-9]+\.[01-9]+)([^01-9a-z]+)/gi;
  var mdr2 = '^\\[[^\]\\|]+\\|[^\\]]+\\][^01-9]+([01-9]+\.[01-9]+)[^01-9a-zA-Z]+';
  var md = '';
  for (var pi = 0; pi < pa.length; pi++) {
    var pt = pa[pi];
    var ptt = pt.getText();
    if (ptt.match(mdr)) {
      md = mdr.exec(ptt);
      pa[pi].replaceText(mdr2, md[1]);
    }
  }
}

function removeShortStartDates() {
  var doc = DocumentApp.openById(Config.get('ANNOUNCEMENTS_2WEEKS_SUNDAY_ID'));
  var paras = doc.getBody().getParagraphs();
  var re = /^\[[^\]\|]+\|[^\]]+\](\D+\d{1,2}\.\d{1,2}\W+)/gi;//matches: "[ foo | bar ] 05.29 ; " or "[ foo | bar ] << baz 05.29 >>; " capturing the portion after [ ]
  
  for (var p in paras) {
    var text = paras[p].getText();
    var match = new RegExp(re).exec(text);//MUST use a new regexp each check or the next check picks up after the previous (making no match size it skips the one at 0)
    if(match)
      paras[p].replaceText(match[1].replace('\[','\\[').replace('\|','\\|').replace('\]','\\]'), ' ');
  }
}

function modifyDatesInBody() {
  updateDatePrototype();
  Logger.clear();
  var doc = DocumentApp.openById(Config.get('ANNOUNCEMENTS_1WEEK_SUNDAY_ID'));
  var opa = doc.getBody().getParagraphs(); //opa contains event text
  
  
  var draftdoc = DocumentApp.openById(Config.get('ANNOUNCEMENTS_2WEEKS_SUNDAY_ID'));
  
  //get fileName
  var filename = draftdoc.getName();
  //get date from file name
  var date = getDateInName(filename);
  //form matching templates
  var templates = createMatchTemplates(date);
  
  draftdoc.getBody().replaceText('Today', '** ERROR **');

  draftdoc.getBody().replaceText('This Monday',    '** ERROR **');
  draftdoc.getBody().replaceText('This Tuesday',   '** ERROR **');
  draftdoc.getBody().replaceText('This Wednesday', '** ERROR **');
  draftdoc.getBody().replaceText('This Thursday',  '** ERROR **');
  draftdoc.getBody().replaceText('This Friday',    '** ERROR **');
  draftdoc.getBody().replaceText('This Saturday',  '** ERROR **');
  
  draftdoc.getBody().replaceText('Next Sunday',    'Today');
  draftdoc.getBody().replaceText('Next Monday',    'This Monday');
  draftdoc.getBody().replaceText('Next Tuesday',   'This Tuesday');
  draftdoc.getBody().replaceText('Next Wednesday', 'This Wednesday');
  draftdoc.getBody().replaceText('Next Thursday',  'This Thursday');
  draftdoc.getBody().replaceText('Next Friday',    'This Friday');
  draftdoc.getBody().replaceText('Next Saturday',  'This Saturday');
  
  
  for (var t = 0; t < templates.length; t++) {
    draftdoc.getBody().replaceText(templates[t].SearchElement, templates[t].ReplaceElement)
    //draftdoc.getBody().replaceText(templates[t].SearchElement,"REPLACED")
  }
}



//Creating templates to search in the document body and replace with the required Term.
function createMatchTemplates(date) {

  var templateReplaceArr = [' Today ', ' This Monday ', ' This Tuesday ', ' This Wednesday ', ' This Thursday ', ' This Friday ', ' This Saturday ', ' Next Sunday '];
  var templateArr = [];
  for (var i = 0; i <= 7; i++) {
    var tempDate = new Date(date.getFullYear(), date.getMonth(), date.getDate());
    tempDate.setDate(date.getDate() + i);
    var dayname = tempDate.getDayName();
    var replacementStr = ((i == 0) ? " Today " : (i == 7) ? " Next " + dayname : " This " + dayname);
    //for(var j=1;j<=7;j++)
    //{ 
    var template = {};
    var regex = getDateRegex(dayname, tempDate.getMonth(), tempDate.getDate());
    template.SearchElement = regex;
    template.ReplaceElement = replacementStr;
    templateArr.push(template);
  }
  return templateArr;
}

function getDateRegex(dayname, month, date) {
  var regexStr = '';
  var dayRegex = getDayRegexStr(dayname);
  var monthRegex = getMonthRegexStr(month);
  var dateRegex = getDateRegexStr(date);
  //var dateRegex = '(' + date+  ')';
  
  //reference Regex /(Sun|Tues)[- \\/., ]+(Oct|October)[- \\/., ]+(10)/gim
  regexStr = dayRegex + '[- /, ]*' + monthRegex + '[- /, ]*' + dateRegex;
  
  return regexStr;
}

function getDayRegexStr(dayname) {
  switch (dayname) {
    case 'Sunday':
      return '(Sunday|Sunday.|Sun|Sun.| )';
      break;
    case 'Monday':
      return '(Monday|Monday.|Mon|Mon.|Mo|Mo.| )';
      break;
    case 'Tuesday':
      return '(Tuesday|Tuesday.|Tue|Tue.|Tues|Tues.| )';
      break;
    case 'Wednesday':
      return '(Wednesday|Wednesday.|Wed|Wed.| )';
      break;
    case 'Thursday':
      return '(Thursday|Th|Thur|Thu|Thrs|Thrs.|Thurs|Thurs.|Th.|TR|TR.| )';
      break;
    case 'Friday':
      return '(Friday|Friday.|Fri|Fri.|FR|FR.| )';
      break;
    case 'Saturday':
      return '(Saturday|Saturday.|Sat|Sat.| )';
      break;
  }
}

function getMonthRegexStr(month) {
  switch (month) {
    case 0:
      return '(January|January.|Jan|Jan.|1\/)';
      break;
    case 1:
      return '(February|February.|Feb|Feb.|2\/)';
      break;
    case 2:
      return '(March|March.|Mar|Mar.|3\/)';
      break;
    case 3:
      return '(April|April.|Apr|Apr.|4\/)';
      break;
    case 4:
      return '(May|May.|5/)';
      break;
    case 5:
      return '(June|June.|Jun|Jun.|6\/)';
      break;
    case 6:
      return '(July|July.|Jul|Jul.|7\/)';
      break;
    case 7:
      return '(August|August.|Aug|Aug.|8\/)';
      break;
    case 8:
      return '(September|September.|Sep|Sep.|Sept|Sept.|9\/)';
      break;
    case 9:
      return '(October|October.|Oct|Oct.|10/)';
      break;
    case 10:
      return '(November|November.|Nov|Nov.|11\/)';
      break;
    case 11:
      return '(December|December.|Dec|Dec.|12\/)';
      break;
      
  }
}

function getDateRegexStr(date) {
  var dateRegex = '';
  //(19[ ]{1}|(19)+$)
  if (parseInt(date) >= 1 && parseInt(date) <= 9) {
    dateRegex = '(' + date + '[ ]{1}|' + '(' + date + ')+$' + ')';
  } else {
    dateRegex = '(' + date + ')';
  }
  
  return dateRegex;
}

//Takes file name , and returns the date mentioned in it , 
//eg filename 'Copy of [ 10.01 ] Sunday Announcements ' will return Sunday, October 1 2017 00;00:000 
function getDateInName(name) {
  var i1 = name.indexOf('[') + 1;
  var i2 = name.indexOf(']');
  var dateArr = name.substring(i1, i2).trim().split('.');
  // var date = new Date(new Date().getFullYear(),parseInt(dateArr[0])-1,parseInt(dateArr[1]));
  // var month = parseInt(dateArr[0].replace('0',''))-1;
  var month = parseInt(dateArr[0].indexOf('0') == 0 ? dateArr[0].replace('0', '') : dateArr[0]) - 1;
  // var day = parseInt(dateArr[1].replace('0',''));
  var day = parseInt(dateArr[1].indexOf('0') == 0 ? dateArr[1].replace('0', '') : dateArr[1]);
  var date = new Date(new Date().getFullYear(), month, day);
  return date;
}

//Update Javascript Date class to add prototype to get date in our desired format
//1. Sunday, October 8
//2. Sunday, Oct 8
//3. Sun, Oct 8
//4. Sun, October 8
function updateDatePrototype() {
  Date.prototype.getFormattedDate = null;
  Date.prototype.getFormattedDate = function(type) {
    
    var daysLong = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
    var daysShort = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
    
    var monthsLong = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
    var monthsShort = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    
    var mm = this.getMonth() + 1; // getMonth() is zero-based
    var dd = this.getDate();
    var dayno = this.getDay();
    var day = '';
    var month = ''
    
    if (type == 1 || type == 2) {
      day = daysLong[dayno];
    } else {
      day = daysShort[dayno];
    }
    
    if (type == 1 || type == 4) {
      month = monthsLong[mm - 1];
    } else {
      month = monthsShort[mm - 1];
    }
    if (type == 5) {
      month = monthsShort[mm - 1];
      
      var returnDate = [
        month + ' ',
        dd
      ].join('');
      
      return returnDate;
    }
    if (type == 6) {
      month = monthsLong[mm - 1];
      
      var returnDate = [
        month + ' ',
        dd
      ].join('');
      
      return returnDate;
    }
    if (type == 7) {
      month = monthsLong[mm - 1];
      
      var returnDate = [
        (mm) + '/',
        dd
      ].join('');
      
      return returnDate;
    }
    var returnDate = [day + ', ',
                      month + ' ',
                      dd
                     ].join('');
    
    return returnDate;
  };
  
  Date.prototype.getDayName = function() {
    var days = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
    return days[this.getDay()];
  };
  
}

