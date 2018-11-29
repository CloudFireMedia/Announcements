//function makeMenu_announcements_master(){
//  DocumentApp.getUi().createMenu('[ Custom Menu ]')
//  .addItem('Add Week', 'PL.master_insertWeeks')
//  .addItem('Add Month', 'PL.master_insertMonth')
//
//  .addSeparator()
//  .addItem('Archive', 'PL.masterToArchive')
////  .addItem('Setup Archive', 'PL.setup')
//
//  .addSeparator()
//  .addItem('Format Document', 'PL.format_master')
//
//  .addSubMenu(
//    DocumentApp.getUi().createMenu('Formatting...')
//    .addItem('Remove Double Spaces', 'PL.format_doubleSpaceToSingle')
//    .addItem('Remove Empty Paragraphs', 'PL.format_removeEmptyParagraphs')
//    .addItem('Fix Page Breaks (and headings)', 'PL.format_master_fixPageBreaksAndHeadings')
//    .addItem('Format Future Events', 'PL.format_master_formatEventsFuture')
//    .addItem('Format Past Events', 'PL.format_master_formatEventsPast')
//  )
//
//  .addSeparator()
//  .addItem('Automation', 'PL.setupAutomation_master')
//
//  .addSubMenu(
//    DocumentApp.getUi().createMenu('Instructions...')
//    .addItem('Document Overview', 'PL.showInstructions_Document')
//    .addItem('Recurring Content', 'PL.showInstructions_RecurringContent')
//  )
////  .addSeparator()
////  .addSubMenu(
////    DocumentApp.getUi().createMenu('Dev...')
////    .addItem('Update Menu', 'makeMenu')
//////    .addItem('test', 'test')
////  )
//
//  .addToUi();
//
//  ////Disabled! This will be moved to a global config 
//  //check to see if the archive doc has been set, if not, warn and ask the to run Archive Setup
//  //but build the menu first or there won't be one if the script times out waiting on the user to click OK
////  if( ! checkSetup()){
////    var archiveSetupWarning = CacheService.getUserCache().get('archiveSetupWarning');
////    if( ! archiveSetupWarning){
////      DocumentApp.getUi().alert('Setup Required', "\
////Archiving is disabled.\n\
////Run the Archive Setup from the [ Custom Menu ] to enable archiving.\n\
////\n(And since I'm so nice, I won't bug you about it for a couple hours.)\
////", DocumentApp.getUi().ButtonSet.OK);
////      CacheService.getUserCache().put('archiveSetupWarning', true, 2*60*60);
////    }
////  }    
//}

function makeMenu_announcements_twoWeeksHence() {
  DocumentApp.getUi().createMenu('[ Custom Menu ]')  
  .addItem('Invite Staff Sponsors to Comment',                                                         'PL.sendMailFunction')
  .addSeparator()
  .addItem('Rotate Content',                                                                           'PL.rotateContent')//was transferText
  .addSeparator()
  .addItem('Copy This Sunday\'s Service Slides to This Sunday\'s \'Live Announcements Slides\' folder', 'PL.moveSlides')

// Submenu commented out by Chad on 05.13.18. Don't think we need this anymore. Can be deleted.
//  .addSubMenu(
////    DocumentApp.getUi().createMenu('➤ Rotate Content - Step by step')
////    .addItem('1 - Return This week to Master',       'PL.moveThisToMaster')
////    .addItem('2 - Move Next week to This week',      'PL.moveNextToThis')
////    .addItem('3 - Move Draft to Next Week',          'PL.moveDraftToNext')
////    .addItem('4 - Get new Master content for Draft', 'PL.moveMasterToDraft')
////    .addItem('5 - Archive oldest week from Master',  'PL.moveOldestToArchive')
// )
  
  .addSeparator()
  .addItem('Re-order Paragraphs',       'PL.reorderParagraphs')
  .addItem('Remove Short Start Dates',  'PL.removeShortStartDates')
  .addItem('Format',                    'PL.draft_callFunctions')
  .addItem('Populate Empty Paragraphs', 'PL.matchEvents')
  .addItem('Update Long Start Dates',   'PL.modifyDatesInBody')
  .addToUi();
}

function makeMenu_announcements_oneWeekHence(){
  DocumentApp.getUi().createMenu('[ Custom Menu ]')
  .addItem('Format All', 'PL.runAllFormattingFunctions_oneWeekOut')
  .addSeparator()
  .addItem('Format Font', 'PL.formatFont_oneWeek')
  .addItem('Format Paragraphs', 'PL.format_removeEmptyParagraphs')
  .addToUi();
}

function makeMenu_announcements_upcomingWeek(){
  DocumentApp.getUi().createMenu('[ Custom Menu ]')
  .addItem('Format', 'PL.runAllFormattingFunctions_upcomingWeek')
  .addItem('Move Service Slides', 'PL.moveSlides')
  .addItem('Email Staff', 'PL.emailStaff')
  .addToUi();
}

function makeMenu_announcements_archive(){
  DocumentApp.getUi().createMenu('[ Custom Menu ]')
  .addItem('Format', 'PL.format_archive')
  .addToUi();
}

