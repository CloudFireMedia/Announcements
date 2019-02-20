function test_inviteStaffSponsorsToComment() {
  inviteStaffSponsorsToComment_()
}

function test_rotateContent() {
  rotateContent_()
}

function test_getWeekTitles() {
  getWeekTitles()
}

function test_getChildren() {
  var body = DocumentApp.openById('11WLqbxCb_NCNAA9sSxn-SD5sM6Mae0yUYe0w-rj-e58').getBody()
  var numberOfChildren = body.getNumChildren()
  for (var childIndex = 0; childIndex < numberOfChildren; childIndex++) {
    Logger.log(body.getChild(childIndex).getType())
  }
}

function test_removeComments() {

  var DOC_ID = '11WLqbxCb_NCNAA9sSxn-SD5sM6Mae0yUYe0w-rj-e58'
//  var DOC_ID = '1L8myiog-o6puURPuWu7R6UOntPRp4-_xfHcdUTo3Yzc'
//  var COMMENT_ID = 'AAAACjkftLo'
  
//  var nextComment = Drive.Comments.get(DOC_ID, COMMENT_ID); 
//  nextComment = Drive.Comments.remove(DOC_ID, COMMENT_ID); 
  
//  Logger.log('before - status: ' + nextComment.status + ', deleted: ' + nextComment.deleted + ', content: ' + nextComment.content); 

// return

  var comments = Drive.Comments.list(DOC_ID); 

  for (var i = 0; i < comments.items.length; i++) { 
  
    var nextComment = comments.items[i]
  
    Logger.log('before - status: ' + nextComment.status + ', deleted: ' + nextComment.deleted + ', content: ' + nextComment.content); 
    
    if (nextComment.status === 'open') {
      Logger.log('Removing ' + nextComment.content)    
      Drive.Comments.remove(DOC_ID, nextComment.commentId)
    }    
  } 
  
} // test_removeComments()

function test_updateComment() {
  var fileId = '11WLqbxCb_NCNAA9sSxn-SD5sM6Mae0yUYe0w-rj-e58'
  var commentId = 'AAAACkLLiH0'
  var resource = {status: 'resolved'}
  var result = Drive.Comments.update(resource, fileId, commentId)
  return
}

function test_listComments() {

  var DOC_ID = '11WLqbxCb_NCNAA9sSxn-SD5sM6Mae0yUYe0w-rj-e58'
//  var DOC_ID = '1L8myiog-o6puURPuWu7R6UOntPRp4-_xfHcdUTo3Yzc'

  var comments = Drive.Comments.list(DOC_ID); 
  
  if (comments.items[0] === undefined) {
    Logger.log('No comments')
    return
  }
  
  var numberOfComments = comments.items.length
  Logger.log('Number of comments: ' + numberOfComments)
  
  if (comments.items && numberOfComments > 0) { 
    for (var i = 0; i < numberOfComments; i++) { 
      var comment = comments.items[i]; 
     Logger.log(typeof comment.deleted);
//     Logger.log('content: ' + comment.content + ', status: ' + comment.status + ', deleted: ' + comment.deleted); 
//      var modifiedDateString = comment.modifiedDate
//      var modifiedDate = new Date(modifiedDateString)
//      var lastWeek = new Date(new Date().getTime() - (7 * 24 * 60 * 60 * 1000))
//      if (modifiedDate > lastWeek) {
//        Logger.log('last week')
//      } else if (modifiedDate < lastWeek) {
//        Logger.log('Longer than a week')
//      }
//      Logger.log('modified: ' + modifiedDateString)
    } 
  } else { 
    Logger.log('No comment found.'); 
  } 
}