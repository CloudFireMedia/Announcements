var SCRIPT_NAME = "Announcements"
var SCRIPT_VERSION = "v0.3.1.dev_ajr"

var TEST_GDOC_ID_ = '11WLqbxCb_NCNAA9sSxn-SD5sM6Mae0yUYe0w-rj-e58';

var config = {

  scriptLastRunText: 'Script last run: ',
  scriptLastRunTextLength: 16,

  notRunYetText: 'The script hasn\’t been run before, how many weeks would you like to go back in the comments (to look for staff members to email)',

  announcements: { // Legacy

    placeholder : '{{CONTENT IN ROTATION}}',
    
    // ==========================================================================================================
    //Feel free to change these to fit your needs. ==============================================================
    
    //notificationEmail - who to tell when someone needs to be told (mostly for error reporting in this script)
    //this can have multiple email addresses like so: ['bob@rupholdt.com','siervo@gmail.com']
    notificationEmail : [],
    
    //format - this consists of several format definitions used to format the document's appearance
    format : {
      heading1 : {
        LINE_SPACING :  1.5,
        SPACING_BEFORE : 12,
        SPACING_AFTER : 0,
        FONT_FAMILY : 'Lato',
        FONT_SIZE : 9,
        HORIZONTAL_ALIGNMENT : DocumentApp.HorizontalAlignment.CENTER,
      },
      heading2 : {
        LINE_SPACING :  1.5,
        SPACING_BEFORE : 12,
        SPACING_AFTER : 0,
        FONT_FAMILY : 'Lato',
        FONT_SIZE : 9,
        HORIZONTAL_ALIGNMENT : DocumentApp.HorizontalAlignment.RIGHT,
        ITALIC : true,
      },
      future : {
        LINE_SPACING :  1.5,
        SPACING_BEFORE : 12,
        SPACING_AFTER : 0,
        FONT_FAMILY : 'Lato',
        FONT_SIZE : 9,
        FOREGROUND_COLOR : '#585858',//dark gray ////Chad, why not black?
        BACKGROUND_COLOR : '#ffffff',
        //      BACKGROUND_COLOR : '#ffddfa',//purplishpink //dev
      },
      futureTitle : {
        BOLD : 'true',
        BACKGROUND_COLOR : '#ffffff',
        //      BACKGROUND_COLOR : '#00ff00',//green //dev
      },
      past : {
        LINE_SPACING :  1.5,
        SPACING_AFTER : 0,
        FONT_FAMILY : 'Lato',
        FONT_SIZE : 9,
        FOREGROUND_COLOR : '#aeaeae',//light gray
        BACKGROUND_COLOR : '#ffffff',
        //      BACKGROUND_COLOR : '#00ffff',//aqua //dev
      },
      archive : {
        FONT_FAMILY : 'Lato',
        FONT_SIZE : 9,
        FOREGROUND_COLOR : '#585858',
        LINE_SPACING : 1.5,
        SPACING_AFTER : 0,
        BOLD : false,
        BACKGROUND_COLOR : '#ffffff',
      },
      
      subtitle : {//the 'nth sunday' text
        HORIZONTAL_ALIGNMENT : DocumentApp.HorizontalAlignment.RIGHT,
        FONT_SIZE      : 9,
        FONT_FAMILY    : 'Lato',
        LINE_SPACING   : 1.5,
        SPACING_BEFORE : 10,
        SPACING_AFTER  : 10,
        ITALIC         : true,
      },
  
      title : {//the 'nth sunday' text
        HORIZONTAL_ALIGNMENT : DocumentApp.HorizontalAlignment.CENTER,
        FONT_SIZE      : 9,
        FONT_FAMILY    : 'Lato',
        LINE_SPACING   : 1.5,
        SPACING_BEFORE : 12,
        SPACING_AFTER  : 0,
        ITALIC         : false,
        BOLD : 'true',
        FOREGROUND_COLOR : '#000000',
      },
    },
    
    // ==========================================================================================================
    //These should probably not be changed. =====================================================================
    
    //sundayPageRegEx - This is how we find the Sunday page titles.  Like [ 04.22 ]
    //if you change this, also change the sundayPagePattern to match
    sundayPageRegEx : /\[ *\d{2}\.\d{2} *]/,
    
    //sundayPagePattern - Same as sundayPageRegEx but used by a different function which can not accept the regular regex version.
    //if you change this, also change the sundayPageRegEx to match
    sundayPagePattern : '\\[ *\\d{2}\\.\\d{2} *] *Sunday *Announcements',//'[ 03.11 ] Sunday Announcements' - must double-escape reserved characters
  //  config.announcements.sundayPagePattern = escapeGasRegExString(config.announcements.sundayPageRegEx);//place this after the config definition
  }  
};