var Comments_ = (function(ns) {

  /**
   * Create a new comment
   */

  ns.create = function(text, fileId) {
  
    if (fileId === undefined) {
      var activeDoc = DocumentApp.getActiveDocument()
      fileId = (activeDoc === null) ? config.testDocId : activeDoc.getId()
    }
    
    var resource = {content: text + new Date()};
    Drive.Comments.insert(resource, fileId)
    
  } // Comments_.create()

  /**
   * Update an existing comment or create a new one
   */

  ns.update = function(text) {

    var activeDoc = DocumentApp.getActiveDocument()
    var fileId = (activeDoc === null) ? config.testDocId : activeDoc.getId()
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
   * @return {Date} the time after which comments are defined as "new"
   */
  
  ns.getLastTimeScriptRun = function(documentId) {
    
    var datetime = null;
    var comments = Drive.Comments.list(documentId);
    var lastTimeInvitesSent = null;
    var lastTimeContentRotated = null;
   
    // Search all the comments
    
    if (comments.items && comments.items.length > 0) {
      
      comments.items.some(function(comment) {
              
        var content = comment.content;
        
        if (content.indexOf(config.lastTimeInviteRunText) !== -1) {
        
          lastTimeInvitesSent = new Date(content.slice(config.lastTimeInviteRunTextLength));
          
        } else if (content.indexOf(config.lastTimeContentRotated) !== -1) {
        
          lastTimeContentRotated = new Date(content.slice(config.lastTimeRotateRunTextLength));
        }        
      })
    }  
    
    // Determine the date/time to return

    if (lastTimeInvitesSent !== null) {    
    
      datetime = lastTimeInvitesSent;
      
    } else if (lastTimeContentRotated !== null) {
    
      datetime = lastTimeContentRotated;
      
    } else {
    
      datetime = new Date((new Date()).getTime() - ONE_WEEK_IN_MS); 
    }

    log('datetime: ' + datetime);
    return datetime;
    
  } // Comments_.getLastTimeScriptRun()
  
  return ns

})(Comments_ || {})