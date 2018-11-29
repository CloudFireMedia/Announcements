function setup() {
  var setup = checkSetup();
  var ui = DocumentApp.getUi();
  var rerunSetup = false;
  
  if(setup){//ok, setup was completed before, can we acces the doc?
    var archiveDocID = config.files.announcements.archive;
    try{ var archiveDoc = DocumentApp.openById(archiveDocID); }catch(e){};

    if( ! archiveDoc){//oh no, we can't use the archive doc, either it's not real, it's gone, or we don't have permission
      var ok = ui.alert('Setup', "Oops! I can't access the archive document.\n\
Make sure it exists, isn't deleted, and the script owner has permission to access it.\n\
Because I can't access it, all I know is the current document ID: "+archiveDocID+" \n\n\
Do you want to change the archive document?\
", ui.ButtonSet.YES_NO_CANCEL);
      if(ok!='YES') return;
      rerunSetup = true;
    }
    
    if( ! rerunSetup){//then show a summary and prompt to replace
      var archiveDocName = archiveDoc.getName();
      var archiveDocURL = archiveDoc.getUrl();
      var ok = ui.alert('Setup', 
                        Utilities.formatString("The archive document is configured as:\n%s\n%s\n\nDo you want to change it?", archiveDocName, archiveDocURL), 
                        ui.ButtonSet.YES_NO_CANCEL);
      if(ok!='YES') return;
      rerunSetup = true;
    }
  
  }
  
  if( ! setup || rerunSetup){//not setup, prompt for archive doc
    var archiveDocInput = ui.prompt('Setup', "\
"+(rerunSetup ? 'Replacing archive doc.':'No archive doc is configured.')+"\n\n\
Please enter the ID or URL of document to use for archiving older announcements.\
", ui.ButtonSet.OK_CANCEL);
    if( ! archiveDocInput.getResponseText() ) return;
    
    var archiveDocID = getIdFromUrl(archiveDocInput.getResponseText());
//    Logger.log(archiveDocInput.getResponseText())
//    Logger.log(archiveDocID)
    PropertiesService.getDocumentProperties().setProperty('archiveDocID', archiveDocID);
    config.files.announcements.archive = archiveDocID;
    
    ui.alert('Setup', "Setup is complete.\n\nTo verify all is well, run this again.", ui.ButtonSet.OK);
  }
}

function checkSetup() {
  var keys = PropertiesService.getDocumentProperties().getKeys();
  if(keys.indexOf('archiveDocID') > -1) return true;
  return false;
}
