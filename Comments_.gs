var Comments_ = (function(ns) {

  /**
   * Create a new comment
   */

  ns.create = function(text, fileId) {
  
    if (fileId === undefined) {
      var activeDoc = DocumentApp.getActiveDocument()
      fileId = (activeDoc === null) ? TEST_DOC_ID_ : activeDoc.getId()
    }
    
    var resource = {content: text + new Date()};
    Drive.Comments.insert(resource, fileId)
    
  } // Comments_.create()

  /**
   * Update an existing comment or create a new one
   */

  ns.update = function(text) {

    var activeDoc = DocumentApp.getActiveDocument()
    var fileId = (activeDoc === null) ? TEST_DOC_ID_ : activeDoc.getId()
    var comments = Drive.Comments.list(fileId);
    var foundComment = false
    
    if (comments.items && comments.items.length > 0) {    
      foundComment = comments.items.some(function(comment) {   
        var content = comment.content
        if (content.indexOf(text) !== -1) { 
          var resource = {content: text + new Date()};
          Drive.Comments.patch(resource, fileId, comment.commentId);
          return true;
        }
      })
    }  
    
    if (!foundComment) {    
      var resource = {content: text + new Date()};
      Drive.Comments.insert(resource, fileId)
    }
    
  } // Comments_.update()
  
  /**
   * Search all the comments for the last time invite for comments were sent and 
   * when the content was last rotated. Based on the last time these functions 
   * were run return the time after which comments are defined as "new".   
   *
   * @return {Date} the time after which comments are defined as "new" or null
   */
   
//    // “hidden” comment is used to store the last time rotateContent() ran
//    WHEN “invite” runs
//        // First get the value to use as last run time
//        IF there is a “last rotate ran time” comment from today
//            Use last time rotate ran
//        ELSE 
//            STOP and ask user to run “rotate” first
//        END IF
//        // Then process each comment
//        FOR EACH comment
//            IF comment is not deleted AND comment is “open” AND comment has been modified since last run AND comment is not the “last time rotate ran”
//                Search comment content for staff names 
//            END IF
//        END FOR EACH
//        // Send one email to each staff member mentioned …
//    END WHEN
   
  ns.getLastTimeScriptRun = function(documentId) {
    
    var comments = Drive.Comments.list(documentId);    
    var lastTimeContentRotated = null;
    var ui = getUi_();
    var title = 'Failed to get last time script run';
    var message
   
    // Search all the comments
    
    if (comments.items && comments.items.length > 0) {      
      comments.items.some(function(comment) {    
        if (comment.status !== "open" || comment.deleted) {
          return false;
        }
        var content = comment.content;      
        if (content.indexOf(config.lastTimeRotateRunText) === 0) {
          lastTimeContentRotated = new Date(content.slice(config.lastTimeRotateRunTextLength));
          return true;
        }        
      })      
    }

    // Determine the date/time to return

    log('getLastTimeScriptRun datetime: ' + lastTimeContentRotated);

    if (lastTimeContentRotated === null) {   
    
      message = 'The content has to be rotated before the invitations to comment can be sent.';
        
      if (ui !== null) {   
        ui.alert(title, message, ui.ButtonSet.OK);
        return null
      } else {
        throw new Error('message');
      }
      
    } else {    
    
      var todayInMs = (new Date()).getTime();
      var ONE_DAY_IN_MS = 24 * 60 * 60 * 1000;
      var yesterdayInMs = todayInMs - ONE_DAY_IN_MS

      if (lastTimeContentRotated.getTime() < yesterdayInMs) {

        message = 'The content was rotate more than a day ago.';

        if (ui !== null) {   
          ui.alert(title, message, ui.ButtonSet.OK);
          return null
        } else {
          throw new Error('message');
        }
      }
     
      return lastTimeContentRotated;
    }
    
  } // Comments_.getLastTimeScriptRun()
  
  return ns

})(Comments_ || {})