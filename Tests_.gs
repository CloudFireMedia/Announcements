// var TEST_DOC_ID_ = '1m2QHvYcC-a9ywnHu7KHHBG7u5ksF2da9uQQwHXP6MaI' // [ 07.21 ] Sunday Announcements - Draft Document
var TEST_DOC_ID_ = '1nwqLAYy9MI6xhzgciSmF5S2m0Aj-ec3PMN8WRM0zY-Y' // Live CCN Week 2 GDoc

function test_misc() {
  while(true)
  var a = 1
}

function test_compareStrings() {
  var a = 'Title_03_17'
  var b = 'Title_03_17'
  var c = compareStrings(a,b)
  return
}

function test_Comments_getLastTimeScriptRun() {
  var a = Comments_.getLastTimeScriptRun(TEST_DOC_ID_)
  return
}

function test_uniq() {
  var a = [1,2,1,3,2,4]
  var b = arrayUnique(a)
  return
}

function test_modifyDatesInBody() {
  modifyDatesInBody_()
  return
}

function test_formatGDoc() {
  formatGDoc_()
  return
}

function test_short() {
  var text = '[ title | FN1 LN1 ] 01.01;'
  var re = /^\[[^\]\|]+\|[^\]]+\](\D+\d{1,2}\.\d{1,2}\W+)/gi;
  var match = new RegExp(re).exec(text);
  return
}

function test_reorderParagraphs() {
  reorderParagraphs_()
}

function test_copySlides() {
  copySlides_()
}

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
  var body = DocumentApp.openById(TEST_DOC_ID_).getBody()
  var numberOfChildren = body.getNumChildren()
  for (var childIndex = 0; childIndex < numberOfChildren; childIndex++) {
    Logger.log(body.getChild(childIndex).getType())
  }
}

// Comments
// --------

function test_createComment() {
  createComment('Test1')
  test_listComments()
}

function test_updateComment() {
  Comments_.update(config.lastTimeRotateRunText);
//  var commentId = 'AAAACkLLiH0'
//  var resource = {status: 'resolved'}
//  var result = Drive.Comments.update(resource, TEST_DOC_ID_, commentId)
  return
}

function test_listComments() {

//  var comments = Drive.Comments.list(TEST_DOC_ID_, {"pageSize": 99}); // NO
//  var comments = Drive.Comments.list(TEST_DOC_ID_, {"pageSize": 100}); // NO
//  var comments = Drive.Comments.list(TEST_DOC_ID_, {pageSize: 100}); // NO
//  var comments = Drive.Comments.list(TEST_DOC_ID_, {fields: '*', pageSize: 100}); // NO
//  var comments = Drive.Comments.list(fileId, optionalArgs)

  var pageToken = ''
  var count = 0

  for (var j = 0; j < 100; j++) {

    var comments = Drive.Comments.list(TEST_DOC_ID_, {fields: '*', pageToken: pageToken}); 
    var pageToken = comments.nextPageToken
    
    if (comments.items[0] === undefined) {
      Logger.log('No comments')
      return
    }
    
    var numberOfComments = comments.items.length
    
    Logger.log('Number of comments: ' + numberOfComments)

    count += numberOfComments

    if (numberOfComments < 20) {
      break;
    }    

    if (comments.items && numberOfComments > 0) { 
    
      for (var i = 0; i < numberOfComments; i++) { 
      
        var comment = comments.items[i]; 
        
        if (comment.content.indexOf(config.lastTimeRotateRunText) === 0) {     
//        Logger.log('content: ' + comment.content + ', status: ' + comment.status + ', deleted: ' + comment.deleted + ', modified: ' + comment.modifiedDate); 
          Logger.log('content: ' + comment.content); 
        }
      } 
      
    } else { 
    
      Logger.log('No comment found.'); 
    }
    
  } // for each call
  
  Logger.log(count)
  
  return
}

function test_getAllComments() {
  var a = getAllComments(TEST_DOC_ID_)
  var b = a.items.length
  return
}