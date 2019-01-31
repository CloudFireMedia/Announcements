/******************************************************************************

Recurring content directives should be enclosed in two sets of angle brackets:
<< first, last, 04.22, before fifth >>

Multiple directives can be included and each will be processed.
Directives can duplicate content so use care.
e.g.: << fifth, last >> will give you two entries if there is a fifth Sunday

Only the first set of options in angle brackets (<< >>) is processed. 
Additional sets are considered part of the content.

Commas are optional but help with readability.
Extra spaces are ignored.

Valid options are:
<< first, second, third, fourth, fifth, last >>
Shown when the week ordinal matches the directive.
"Fifth" will not be shown if the month has no fifth Sunday.  Use "last" for this.
"Last" will be shown on the Fourth or Fifth as appropriate.

<< before fifth >>
Shown on the fourth Sunday only when there is a fifth Sunday in the month.

<< 04.22, 10.28 >>
Shown on the exact date only.
If the date given is not a Sunday, it will not be shown.
Note that the format requires the leading zero if the month or day is a single digit.

*****************************************************************************

*/

// Public API
// ==========

// Master Sunday
// -------------

function master_insertMonth()                {master_insertMonth_()}
function master_insertWeeks()                {master_insertWeeks_()}
function format_master()                     {format_master_()}
function showInstructions_Document()         {showInstructions_Document_()}
function showInstructions_RecurringContent() {showInstructions_RecurringContent_()}

// This (upcoming) Sunday (0 weeks)
// --------------------------------

function runAllFormattingFunctions_upcomingWeek() {runAllFormattingFunctions_upcomingWeek_()}
function moveSlides()                             {moveSlides_()}
function emailStaff()                             {emailStaff_()}
function emailStaff_submit(executiveAssitant)     {emailStaff_submit_(executiveAssitant)}

// Next Sunday (1 weeks)
// ---------------------

function runAllFormattingFunctions_oneWeekOut() {runAllFormattingFunctions_oneWeekOut_()}
function formatFont_oneWeek()                   {formatFont_oneWeek_()}

// Draft Sunday (2 weeks)
// ----------------------

function inviteStaffSponsorsToComment()     {inviteStaffSponsorsToComment_()}
function rotateContent()                    {rotateContent_()}
// function moveSlides()                       {moveSlides_()} // 0 weeks
function reorderParagraphs()                {reorderParagraphs_()}
function removeShortStartDates()            {removeShortStartDates_()}
function draft_callFunctions()              {draft_callFunctions_()}
function matchEvents()                      {matchEvents_()}
function modifyDatesInBody()                {modifyDatesInBody_()}
function countInstancesofLiveAnnouncement() {countInstancesofLiveAnnouncement_()}
function cleanInstancesofLiveAnnouncement() {cleanInstancesofLiveAnnouncement_()}

// Archive
// -------

function formatFont_archive() {formatFont_archive_()}