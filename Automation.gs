/*

insertWeeks
masterToArchive

*/

//function setupAutomation_master(){
//
//  if( ! isAuthorizedToRun()){
//    //deny access if not authorized
//    DocumentApp.getUi().alert('Automation: Unauthorized', "Oops!  You're not authorized to run this.\n\nOnly the document owner can enable automation.", DocumentApp.getUi().ButtonSet.OK)
//    return;
//  }
//  
//  var insertTrigger = hasTrigger('insertWeeks');//or insertMonth
//  var archiveTrigger = hasTrigger('masterToArchive');
//  var enabled = insertTrigger && archiveTrigger;
//
//  if(DocumentApp.getUi().alert('Automation', "\
//Automation is currently "+(enabled ? 'ON':'OFF')+".\n\
//This will "+(enabled ? 'DISABLE':'ENABLE')+" the automatic creation and archiving of pages.\n\
//\nContinue?", DocumentApp.getUi().ButtonSet.YES_NO_CANCEL)
//       !='YES') return;//stop if they don't click the YES button
//  
//  //delete existing triggers in either case
//  //this both ensures only a single trigger is installed
//  //and that the running trigger has the most recent preferences
//  deleteTriggerByHandlerName('insertWeeks');
//  deleteTriggerByHandlerName('masterToArchive');
//  
//  if( ! enabled){
//    ScriptApp.newTrigger('insertWeeks').timeBased().everyWeeks(1).onWeekDay(ScriptApp.WeekDay.SUNDAY).atHour(3).create()
//    ScriptApp.newTrigger('masterToArchive').timeBased().everyWeeks(1).onWeekDay(ScriptApp.WeekDay.SUNDAY).atHour(3).create()
//  }
//  
//  DocumentApp.getUi().alert('Automation', "Automation has been "+(enabled ? 'DISABLED':'ENABLED')+".", DocumentApp.getUi().ButtonSet.OK);
//  
//  // Private Functions
//  // -----------------
//
//  function isAuthorizedToRun(){
//  
//    ///this is left from the standalone version. It would be ebtter to rewrite it to work as a helper
//    ///although, as of now this is the only 'module' that uses it
//    
//    
//    //test current user to see if in list of authorized users to run
//    //this goes in the first line of the trigger itself
//  
//    // Not using this portion in this script, just checking the owner
//    //  PropertiesService.getScriptProperties().setProperty('authorizedAccounts',JSON.stringify(['admin@example.com']));
//  
//    //get authorized user array
//    var authorizedAccounts = JSON.parse(PropertiesService.getScriptProperties().getProperty('authorizedAccounts')) || [];
//    //add document owner - this way the owner can change without adding/removing from the auth list
//    //if it's a standalone script, use ScriptApp.getScriptId() instead of the bound spreadsheet id
//    authorizedAccounts.push(DriveApp.getFileById(DocumentApp.getActiveDocument().getId()).getOwner().getEmail().toLowerCase());//this will break in team drives
//  
//    if(authorizedAccounts.indexOf(Session.getEffectiveUser().getEmail().toLowerCase()) > -1) return true;
//    //else
//    //send email about unauthorized attempt to run
//    var subject = 'Unauthorized attempt to run the '+DocumentApp.getActiveDocument().getName()+' script made on '+Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yy 'at' hh:mm ZZZ");
//    var subject = Utilities.formatString('Unauthorized attempt to run the %s script made on ', DocumentApp.getActiveDocument().getName(), fDate(null, "dd/MM/yy 'at' hh:mm ZZZ"));
//    var body = Utilities.formatString('The script %s was run by %s who is not in the authorized users list.<br><br>Authorized accounts are:<br>', DocumentApp.getActiveDocument().getName(), Session.getActiveUser());
//    for(var i=0;i<authorizedAccounts.length;i++)
//      body += '&nbsp;&nbsp;&nbsp;&nbsp;'+authorizedAccounts[i]+'<br>'
//    
//    MailApp.sendEmail({to:config.notificationEmail, subject:subject, htmlBody:body});
//    return false;
//  }
//  
//} // setupAutomation_master()