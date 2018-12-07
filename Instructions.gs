function showInstructions(html, title){
  var modal = HtmlService.createHtmlOutput(html).setWidth(800).setHeight(600);
  DocumentApp.getUi().showModalDialog(modal, title );
}

function showInstructions_COPYMEFORNEWINSTRUCTIONSBLOCK(){
  var title = 'Title'
  var html = "<h1>Instruction</h1> \
<h2>"+title+"</h2> \
\
Brief Description<br>\
<br>\
<h3>Sub-title</h3>\
\
<dl>\
\
<dt>Something</dt>\
<dd>\
Words about something\
<br><br>\
</dd>\
\
</dl>\
";
  
  showInstructions(html, title);
  
}

function showInstructions_Document_(){
  var title ='Document Overview'; 
  var html = "<h1>"+title+"</h1> \
\
\
This is the Master Document for Sunday Announcements. \
Gold, Silver and Bronze submissions are written to this document \
by their respective forms.<br>\
<br>\
<h3>Distinct actions performed by this script are:</h3>\
<dl>\
<dt>Add Week</dt>\
<dd>Adds a page at the beginning for the week following the first Sunday announcement page.<br><br></dd>\
\
<dt>Add Month</dt>\
<dd>Same as Add Week but adds weeks up to the end of current month.<br>\
If the first added week is a new month, adds the entire new month.<br><br></dd>\
\
<dt>Archive</dt>\
<dd>Moves old pages to the document specified during setup<br><br></dd>\
\
<dt>Archive Setup</dt>\
<dd>Required to use the Archive function.\
The user is prompted for the ID or URL of the document to save archives to.<br>\
On document open, if the archive setup has not been run, the user is prompted to do so.<br>\
Once so notified, that user will not be prompted again for another two hours.<br>\
(And then, only if they reopen the document.)<br><br></dd>\
\
<dt>Format Document</dt>\
<dd>Corrects document formatting by:<br>\
Removing double spaces<br>\
Removing empty paragraphs<br>\
Correcting page breaks where missing or doubled<br>\
Settings page title headings and subheadings<br>\
Standardizing font attributes<br>\
Adjusting color for future and past pages<br>\
</ul>\
<br/>\
<dt>Automation</dt>\
<dd>Enables the automatic insertion and archiving of pages each week.<br><br></dd>\
";
  
  showInstructions(html, title);
}

function showInstructions_RecurringContent_(){
  var title = 'Recurring Content'
  var html = "<h1>"+title+"</h1> \
\
Recurring Content is automatically added to each new Sunday page as it is created.  \
Recurring Content directives are added to the page entitles, aptly enough, [ RECURRING CONTENT ]<br>\
<br>\
<h3>Directives</h3>\
Recurring content directives should be enclosed in two sets of angle brackets:<br>\
<< first, last, 04.22, before fifth >><br>\
<br>\
Multiple directives can be included and each will be processed.<br>\
Directives can duplicate content so use care.<br>\
e.g.: << fifth, last >> will give you two entries if there is a fifth Sunday<br>\
<br>\
Only the first set of options in angle brackets (<< >>) is processed. <br>\
Additional sets are considered part of the content.<br>\
<br>\
Commas are optional but help with readability.<br>\
Extra spaces are ignored.<br>\
\
<h3>Valid Options</h3>\
<dl>\
\
<dt>Ordinals</dt>\
<dd>\
Example: << first, second, third, fourth, fifth, last >><br>\
Shown when the week ordinal matches the directive.<br>\
\"Fifth\" will not be shown if the month has no fifth Sunday.  Use \"last\" for this.<br>\
\"Last\" will be shown on the Fourth or Fifth as appropriate.<br>\
<br><br>\
</dd>\
\
<dt>Sunday Before Fifth Sunday</dt>\
<dd>\
Example: << before fifth >><br>\
Shown on the fourth Sunday only when there is a fifth Sunday in the month.<br>\
<br><br>\
</dd>\
\
<dt>Specified Date</dt>\
<dd>\
Example: << 04.22, 10.28 >><br>\
Shown on the exact date only.<br>\
If the date given is not a Sunday, it will not be shown.<br>\
Note that the format requires the leading zero if the month or day is a single digit.<br>\
<br><br>\
</dd>\
\
</dl>\
";
  
  showInstructions(html, title);
  
}

