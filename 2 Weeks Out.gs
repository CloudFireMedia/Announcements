/**
 * A high level function that does various processing of the 2 week GDoc
 */

function process2WeekDoc_() {
  rotateContent_();
  copySlides_();
  formatGDoc_();
  reorderParagraphs_();
  removeShortStartDates_();
  updateWeek2EventDescriptions_();
  formatGDoc_();
  modifyDatesInBody_();
  countInstancesofLiveAnnouncement_();
  formatGDoc_();
}

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

function formatGDoc_() {

  var doc = DocumentApp.openById(Config.get('ANNOUNCEMENTS_2WEEKS_SUNDAY_ID'));
  var body = doc.getBody();
  var paras = body.getParagraphs();
  var length = paras.length;

  removeRowReferences();
  mergeParagraph();
  formatParagraph();
  boldBetweenSquareBrackets();
  highlightStaffSponsorNames();
  
  return;
  
  // Private Functions
  // -----------------
  
  /**
   * Remove the row reference from the event title: [ title | row ref | sponsor name ]
   */
   
  // TODO - Although the filtering is done the new title is never written back to the 
  // GDoc - https://trello.com/c/orpc2fz4/534-removerowreference-doesnt
   
  function removeRowReferences() {
  
    for (i = 0; i < length; ++i) {
    
      var str = paras[i].getText();
      
      if (str.indexOf("of the month") === -1) {
      
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
            if (str.indexOf("]") !== -1) {
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
    
  } // formatGDoc_.removeRowReferences()

  /**
   * Tidy up the CRs and spaces around text after ">>" in the event description. If the >> at the 
   * end of the event description is within 5 characters of the left hand margin, remove any 
   * CRs between it and the previous event text, along with any spaces to the left of it. 
   */

  function mergeParagraph() {
  
    var textLocation = {};
    var i;
    var cnt = 0;
    
    for (i = 0; i < length; ++i) {
    
      var paraText = paras[i].getText();
      
      if (paraText != "") {
      
        if (cnt != 0) {
        
          paraText = paraText.trim();
          
          if (paraText.indexOf(">>") == 0 || 
              paraText.indexOf(">>") == 1 || 
              paraText.indexOf(">>") == 2 || 
              paraText.indexOf(">>") == 3 ||
              paraText.indexOf(">>") == 4) {
          
            var paraTextPrevious = paras[i - 1].getText();
            paraTextPrevious = paraTextPrevious.replace(/[\r\n]/g, "");
            paraText = paraText.replace(/[\r\n]/g, "");
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
    
  } // formatGDoc_.mergeParagraph()
  
  /**
   * Set headings and paragraphs to standard formatting.
   */
  
  function formatParagraph() {
  
    var cnt = 0;
    
    for (i = 0; i < length; ++i) {
    
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
    
  } // formatGDoc_.formatParagraph()
  
  function boldBetweenSquareBrackets() {
  
    var startTag = "[\[]";
    var endTag = "[\]]";
    var i = 0;
    
    for (i = 0; i < length; ++i) {
    
      var from = paras[i].findText(startTag);
      var to = paras[i].findText(endTag, from);
      
      if ((to != null && from != null) && ((to.getStartOffset()) - (from.getStartOffset() + startTag.length) > 0)) {
        paras[i].editAsText().setBold(from.getStartOffset(), to.getStartOffset(), true);
      }
    }
    
  } // formatGDoc_.boldBetweenSquareBrackets()
  
  function highlightStaffSponsorNames() {
  
    var textToHighlight = checkStaffDataSheet();
    var highlightStyle = {};
    highlightStyle[DocumentApp.Attribute.BACKGROUND_COLOR] = '#FCFC00';
    highlightStyle[DocumentApp.Attribute.BOLD] = true;
    var textLocation = {};
    var i;
    
    for (i = 0; i < length; ++i) {
    
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
      
    } // formatGDoc_.highlightStaffSponsorNames.checkStaffDataSheet()
    
  } // formatGDoc_.highlightStaffSponsorNames()
    
} // formatGDoc_()

/**
 * Search the document and comment text for staff to email for comments
 */

/*
// Two separate “hidden” comments will be used to store the last time
// rotateContent() ran and inviteComments() ran

WHEN “invite” runs
  // First get the value to use as last run time
  IF there is a “last time invite ran” comment
    Use last time invite ran
  ELSE 
    IF there is a “last rotate ran time” comment
      Use last time rotate ran
    ELSE
      Use 1 week ago as last run time
    END IF
  END IF
  // Then process each comment
  FOR EACH comment
    IF comment is not deleted AND comment is “open” AND comment has been modified 
      since last run AND comment is not the “last time invite ran” AND comment is not the “last time rotate ran”
      Search comment content for staff names 
    END IF
  END FOR EACH
  // Send one email to each staff member mentioned ...
END WHEN

*/


function inviteStaffSponsorsToComment_() {

  var docSunday = DocumentApp.getActiveDocument();
  
  if (docSunday === null) {
    docSunday = DocumentApp.openById(TEST_DOC_ID_);
  }

  var documentId = docSunday.getId();
  var arrnew = docSunday.getName().split("-");
  var paraFirst = arrnew[0];
  var str = arrnew[0];
  var arr = str.split("[");
  var arr1 = arr[1].split("]");
  var documentShortDate = arr1[0].trim();
  checkName(documentShortDate);
  cleanInstancesofLiveAnnouncement_();
  
  return;
  
  // Private Functions
  // -----------------

  function checkName(documentShortDate) {
  
    var staffToEmail = makeStaffMailList();
    var rawCommentData = Drive.Comments.list(documentId);
    var commentsContent = getOpenCommentsContent();
    var emailListArray = [];
    
    // Make a list of all the staff mentioned in text and comments
    
    staffToEmail.forEach(function(emailAddress) {
    
      var staffName = emailAddress[0];
      var staffEmail = emailAddress[1];
      
      if (staffEmail !== "" && (foundNameInText(staffName) || foundNameInComments(staffName))) { 
        emailListArray.push(staffEmail);
      }      
    })
    
    // Filter the list to one mention of each
    
    var emailList = "";
    
    getUniqueItems(emailListArray).forEach(function(nextEmail) {    
      emailList += nextEmail +  ",";
    })
    
    // Send each an email    
    sendDraftMailFinal(emailList, documentShortDate);
    
    Comments_.update(config.lastTimeInviteRunText);
    
    return;
    
    // Private Functions
    // -----------------

    function getUniqueItems(a) {
      var seen = {};
      var newArray = a.filter(function(item) {
        return seen.hasOwnProperty(item) ? false : (seen[item] = true);
      })
      return newArray;
    }


    function foundNameInComments(staffName) {
      return commentsContent.indexOf(('' + staffName).toUpperCase()) !== -1
    }

    function foundNameInText(staffName) {
    
      var check = 0;
      staffName = staffName.trim();
      
      if (staffName !== "") {
      
        var body = docSunday.getBody();
        var textToHighlight = staffName;
        var paras = body.getParagraphs();
        var textLocation = {};
        var i;
        var result = "";
        
        for (i = 0; i < paras.length; ++i) {
        
          textLocation = paras[i].findText(textToHighlight);
          
          if (textLocation != null && textLocation.getStartOffset() != -1) {
            return true;
          }
        }
      }
      
      return false;
    }

    function sendDraftMailFinal(emailList, documentShortDate) {
              
      if (emailList === '') {
        log('No emails sent')
        return;
      }
      
      var sundayAnnouncementsDraftDocumentUrl = DocumentApp.openById(documentId).getUrl();
      var subject = Utilities.formatString("Please Review: [ %s ] Sunday Announcements draft", documentShortDate);
      var body = Utilities.formatString("Dear Event Sponsor: <br><br> \
    Please review the document linked below regarding promotion for your upcoming event.<br><br>\
    You are invited to suggest changes to your event's promotion by typing directly in the document by Friday EOD. <br><br>\
    Thank you!<br><br>--<br>\
    <a href='%s'>[ %s ] Sunday Announcement draft</a>", 
                                        sundayAnnouncementsDraftDocumentUrl,
                                        documentShortDate
                                       );
      emailList = emailList.replace(/\,$/, ''); //should build the list without the trailing comma in the first place
      
      MailApp.sendEmail({
        to: emailList,
        subject: subject,
        htmlBody: body
      });
            
      log('Emails sent to ' + emailList)
      
    } // sendDraftMailFinal()

    /**
     * Get comments content if it passes the various criteria described, mainly
     * only since the script was last run
     */

    function getOpenCommentsContent() {
    
      var commentsContent = "";
      var lastTimeScriptRun = Comments_.getLastTimeScriptRun(documentId)

      for (var i = 0; i < rawCommentData.items.length; i++) {
      
        var nextComment = rawCommentData.items[i]        
        var modifiedDate = new Date(nextComment.modifiedDate)
                
        if (nextComment.content.indexOf(config.lastTimeInviteRunText) === -1 && 
            nextComment.content.indexOf(config.lastTimeRotateRunText) === -1 &&
            nextComment.status === "open" && 
            !nextComment.deleted && 
            modifiedDate > lastTimeScriptRun) {
        
          // This is not the "script last run" AND it is open AND it hasn't been 
          // deleted AND it was created since the script was last run
          commentsContent += "  " + (nextComment.content)
        }
      }
      
      return ('' + commentsContent).toUpperCase();
      
    } // inviteStaffSponsorsToComment_.checkName.getOpenCommentsContent()
    
  } // inviteStaffSponsorsToComment_.checkName()
  
} // inviteStaffSponsorsToComment_()

function makeStaffMailList() { 
  var staffSheet = SpreadsheetApp.openById(Config.get('STAFF_DATA_GSHEET_ID'));
  var numRows = staffSheet.getLastRow() - 2;
  var staffRows = staffSheet.getSheetValues(3, 1, numRows, 2);
  var staffEmails = staffSheet.getSheetValues(3, 9, numRows, 1);
  var staffFromSheet = [];
  
  for (var row in staffRows) {
    staffFromSheet.push([staffRows[row][0] + ' ' + staffRows[row][1], staffEmails[row][0]]);
  }
  return staffFromSheet;
}

function reorderParagraphs_() {
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
      npa[npi].text = ptt;
      npa[npi].sort = true;
    }
    npi++;
  }
  var tp = '';
  for (npi = 0; npi < npa.length; npi++) {
    if (npa[npi].sort == true) {
      var tpi = npi;
      for (var npi2 = npi + 1; npi2 < npa.length; npi2++) {
        if (npa[npi2].sort == true) {
          if (npa[npi2].md < npa[tpi].md) {
            tpi = npi2;
          }
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
  for (var pi = 0; pi < pa.length; pi++) {
    doc.getBody().removeChild(pa[pi]);
  }
  for (npi = 0; npi < npa.length; npi++) {
    doc.getBody().appendParagraph(npa[npi].paragraph);
  }
  doc.getBody().removeChild(doc.getBody().getParagraphs()[0]);
  
} // reorderParagraphs_()

function updateWeek2EventDescriptions_() {
  
  var doc = DocumentApp.openById(Config.get('ANNOUNCEMENTS_1WEEK_SUNDAY_ID'));
  var opa = doc.getBody().getParagraphs(); //opa contains event text
  
  var draftdoc = DocumentApp.openById(Config.get('ANNOUNCEMENTS_2WEEKS_SUNDAY_ID'));
  
  //will return Paragraph array
  var draftDocParagraphArray = draftdoc.getBody().getParagraphs(); //contains event and style content
  var npa = [];
  var npi = 0;
  var md = '';
  //var mdr=/^\[[^\]\|]+\|[^\]]+\][^01-9]+([01-9]+\.[01-9]+)[^01-9]/gi ;
  //var mdr2=/^\[[^\]\|]+\|[^\]]+\][^01-9]+([01-9]+\.[01-9]+)[^01-9].*[a-z]+/gi ;
  
  //Build regex object
  var mdr = /^\[([^\]\|]+\|[^\]]+)\]/gi; // [ ... | ... ]
  var mdr2 = /^\[[^\]\|]+\|[^\]]+\].*[a-z]+/gi; // [ ... | ... ] ...
  var mdrblank = /[a-z]+/gi;
  
  //DRAFT DOC:  iterate through the paragraph array
  
  draftdoc.getBody().appendParagraph(" ");
  
  for (var pi = 0; pi < draftDocParagraphArray.length; pi++) {
  
    var pt = draftDocParagraphArray[pi];
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
  
} // updateWeek2EventDescriptions_()

function removeShortStartDates_() {

  var doc = DocumentApp.openById(Config.get('ANNOUNCEMENTS_2WEEKS_SUNDAY_ID'));
  var paras = doc.getBody().getParagraphs();
  
  //matches: "[ foo | bar ] 05.29 ; " or "[ foo | bar ] << baz 05.29 >>; " capturing the portion after [ ]
  var re = /^\[[^\]\|]+\|[^\]]+\](\D+\d{1,2}\.\d{1,2}\W+)/gi;
  
  for (var p in paras) {
  
    var text = paras[p].getText();
    
    //MUST use a new regexp each check or the next check picks up after the previous (making no match size it skips the one at 0)
    var match = new RegExp(re).exec(text);
    
    if (match) {
      paras[p].replaceText(match[1].replace('\[','\\[').replace('\|','\\|').replace('\]','\\]'), ' ');
    }
  }
}

function modifyDatesInBody_() {

  updateDatePrototype();
  
  var doc = DocumentApp.openById(Config.get('ANNOUNCEMENTS_1WEEK_SUNDAY_ID'));
  var opa = doc.getBody().getParagraphs(); //opa contains event text  
  var draftdoc = DocumentApp.openById(Config.get('ANNOUNCEMENTS_2WEEKS_SUNDAY_ID'));
  var filename = draftdoc.getName();
  var date = getDateInName(filename);
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
  }
  
  return;
  
  // Private Functions
  // -----------------
  
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
    
  } // modifyDatesInBody_.updateDatePrototype()

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
    
  } // modifyDatesInBody_.createMatchTemplates()

} // modifyDatesInBody_()

function cleanInstancesofLiveAnnouncement_() {

  var twoWeeksId = Config.get('ANNOUNCEMENTS_2WEEKS_SUNDAY_ID')
  var doc = DocumentApp.openById(twoWeeksId);
  var docID = doc.getId();
  var body = doc.getBody();
  body.editAsText().replaceText('[0-9]+\\s✅\\s', "");
  
} // cleanInstancesofLiveAnnouncement_()

function countInstancesofLiveAnnouncement_() {

  cleanInstancesofLiveAnnouncement_()
  
  var regE = new RegExp('\\[\\s([^\\|]*)\\s\\|\\s([^\\]]*)\\s\\]', 'ig');
  
  var twoWeeksId = Config.get('ANNOUNCEMENTS_2WEEKS_SUNDAY_ID')
  var doc = DocumentApp.openById(twoWeeksId);
  var docID = doc.getId();
  var body = doc.getBody();
  var paragraphs = body.getParagraphs();
  
  for (var i = 0; i < paragraphs.length; i++) {
    
    var parag = paragraphs[i]   
    var parString = parag.getText();
    var matches = parString.match(regE);
    var match;
    
    if (match = regE.exec(parString)) {
    
      var event = match[1];
      
      if (('' + event).trim().toUpperCase() != "EVENT NAME") {      
        var counter = 0;
        counter = counter + searchUpcoming(event);
        counter = counter + searchOneWeek(event);
        counter = counter + searchTwoWeeks(event);
        counter = counter + searchMaster(event);
        var counterS = "" + counter + " ✅ ";
        parag.insertText(0, counterS).setBackgroundColor('#ADFF2F')        
      }
    }
  }
  
  return
  
  // Private Functions
  // -----------------
  
  function searchUpcoming(strEvent) {
    
    var result = 0;
    
    var regE = new RegExp('\\[\\s([^\\|]*)\\s\\|\\s([^\\]]*)\\s\\]', 'ig');
      
    var zeroWeeksId = Config.get('ANNOUNCEMENTS_0WEEKS_SUNDAY_ID')    
    var doc = DocumentApp.openById(zeroWeeksId);
    var body = doc.getBody();
    var paragraphs = body.getParagraphs();
    
    for (var i = 0; i < paragraphs.length; i++) {
      
      var parag = paragraphs[i]
      var parString = parag.getText();
      var matches = parString.match(regE);
      
      var match;
      
      if (match = regE.exec(parString)) {
      
        var event = match[1];
        
        if (('' + event).trim().toUpperCase() == ('' + strEvent).trim().toUpperCase()) {
          result++;
        }
      }      
    }
    
    return result
    
  } // countInstancesofLiveAnnouncement_.searchUpcoming()
  
  function searchOneWeek(strEvent) {
    
    var result = 0;
    
    var regE = new RegExp('\\[\\s([^\\|]*)\\s\\|\\s([^\\]]*)\\s\\]', 'ig');

    var oneWeekId = Config.get('ANNOUNCEMENTS_1WEEK_SUNDAY_ID')    
    var doc = DocumentApp.openById(oneWeekId);
    var body = doc.getBody();
    var paragraphs = body.getParagraphs();
    
    for (var i = 0; i < paragraphs.length; i++) {
      
      var parag = paragraphs[i]
      var parString = parag.getText();
      var matches = parString.match(regE);
      
      var match;
      if (match = regE.exec(parString)) {
      
        var event = match[1];
        
        if (('' + event).trim().toUpperCase() == ('' + strEvent).trim().toUpperCase()) {
          result++;
        }
      }
    }
    
    return result
    
  } // countInstancesofLiveAnnouncement_.searchOneWeek()
  
  function searchTwoWeeks(strEvent) {
    
    var result = 0;
    
    var regE = new RegExp('\\[\\s([^\\|]*)\\s\\|\\s([^\\]]*)\\s\\]', 'ig');

    var twoWeeksId = Config.get('ANNOUNCEMENTS_2WEEKS_SUNDAY_ID')
    var doc = DocumentApp.openById(twoWeeksId);
    var body = doc.getBody();
    var paragraphs = body.getParagraphs();
    
    for (var i = 0; i < paragraphs.length; i++) {
      
      var parag = paragraphs[i];
      var parString = parag.getText();
      var matches = parString.match(regE);
      
      var match;
      
      if (match = regE.exec(parString)) {
      
        var event = match[1];
        
        if (('' + event).trim().toUpperCase() == ('' + strEvent).trim().toUpperCase()) {
          result++;
        }
      }      
    }
    
    return result
    
  } // countInstancesofLiveAnnouncement_.searchTwoWeeks()
  
  function searchMaster(strEvent) {
    
    var result = 0;
    
    var regE = new RegExp('\\[\\s([^\\|]*)\\s\\|\\s([^\\]]*)\\s\\]', 'ig');
    var regE1 = new RegExp('\\[\\s([0-9][0-9])\\.([0-9][0-9])\\s\\]\\sSunday\\sAnnouncements', 'ig');

    var masterId = Config.get('ANNOUNCEMENTS_MASTER_SUNDAY_ID')
    var doc = DocumentApp.openById(masterId);
    var body = doc.getBody();
    var paragraphs = body.getParagraphs();
    
    var start = false;
    var end = false;
    var str2Analise = "";
    
    var res;
    var resultF = body.findText("{{CONTENT IN ROTATION}}");
    
    if (resultF != null) {
    
      while (resultF != null) {      
        res = resultF;
        resultF = body.findText("{{CONTENT IN ROTATION}}", resultF)  
      }
      
      var el = res.getElement();
      var parent = el.getParent();
      
      while (parent.getType() != "BODY_SECTION") {
        el = el.getParent();
        parent = el.getParent();
      }
      
      var startIndex = body.getChildIndex(el);
      
      var counter = 0;
      
      for (var i = 0; i < paragraphs.length; i++) {
        
        var parag = paragraphs[i]
        
        if (body.getChildIndex(parag) > startIndex) {
          
          var parString = parag.getText();      
          var matches = parString.match(regE1);
          var match;
          
          if (match = regE1.exec(parString)) {
          
            counter++
              
            if (counter == 4) {
              break;
            }
          }
          
          var matches1 = parString.match(regE);          
          var match1;
          
          if (match1 = regE.exec(parString)) {
          
            var event = match1[1];
            
            if (('' + event).trim().toUpperCase() == ('' + strEvent).trim().toUpperCase()) {
              result++;
            }
          }          
        }
      }
    }
    
    return result
    
  } // countInstancesofLiveAnnouncement_.searchMaster()

} // countInstancesofLiveAnnouncement_()