function getUpcomingSunday(date, skipTodayIfSunday) {
  //return the next Sunday, which might be today
  //skipTodayIfSunday skips this Sunday and returns next week Sunday
  log( '--getUpcomingSunday('+(date ? fDate(date) : 'null')+')' )
  date = new Date(date || new Date());//clone the date so as not to change the original
  date.setHours(0,0,0,0);
  if( skipTodayIfSunday || date.getDay() >0)//if it's not a Sunday...
    date.setDate(date.getDate() -date.getDay() +7);//subtract days to get to Sunday then add a week
  log('upcomingSunday returned: '+fDate(date));
  return date;
}

function log(message) {
  Logger.log(message)
}

//returns the date formatted with format, default to today if date not provided
function fDate(date, format){
  date = date || new Date();
  format = format || "MM/dd/yy";
  return Utilities.formatDate(new Date(date), Session.getScriptTimeZone(), format)
}

/**
* Add time to a date in specified interval
* Negative values work as well
*
* @param {date} javascript datetime object
* @param {interval} text interval name [year|quarter|month|week|day|hour|minute|second]
* @param {units} integer units of interval to add to date
* @return {date object} 
*/

function dateAdd(date, interval, units) {
  date = new Date(date); //don't change original date
  switch(interval.toLowerCase()) {
    case 'year'   :  date.setFullYear(date.getFullYear() + units);            break;
    case 'quarter':  date.setMonth   (date.getMonth()    + units*3);          break;
    case 'month'  :  date.setMonth   (date.getMonth()    + units);            break;
    case 'week'   :  date.setDate    (date.getDate()     + units*7);          break;
    case 'day'    :  date.setDate    (date.getDate()     + units);            break;
    case 'hour'   :  date.setTime    (date.getTime()     + units*60*60*1000); break;
    case 'minute' :  date.setTime    (date.getTime()     + units*60*1000);    break;
    case 'second' :  date.setTime    (date.getTime()     + units*1000);       break;
    default       :  date = undefined; break;
  }
  return date;
}

function getDateFromShortcode(shortcode){//shortcode = '03.10' or contains such
  var monthDay = shortcode.split('.');
  if( ! monthDay.length == 2) throw 'Invalid shortcode';
  var month = monthDay[0]-1;
  var day   = monthDay[1];
  var year  = parseInt(month) < new Date().getMonth()+1 //of month < current month
  ? new Date().getFullYear()+1 // it's next year
  : new Date().getFullYear();  // else it's this year
  
  return new Date(year, month, day);
}

function getLastSundayOfMonth(date){
  date = new Date(date);
  date.setMonth( date.getMonth() +1 );//add a month
  date.setDate(0);//subtract one day for last day of previous month
  date.setDate( date.getDate() - date.getDay() );//subtract current day to get to Sunday prior
  
  return date;
}

function compareStrings(s1, s2) {
  var sa1 = s1.split(' ');
  var sa2 = s2.split(' ');
  var s1cnt = 0;
  var s1l = sa1.length;
  var s1i = 0;
  var sa1tmp = [];
  for (s1i = 0; s1i < s1l; s1i++) {
    if (sa1[s1i].match(/.*[a-z]{2}.*/i)) {
      sa1tmp[s1cnt] = sa1[s1i].replace(/[^a-z]+/i, '');
      s1cnt++;
    }
  }
  sa1 = sa1tmp;
  var s2cnt = 0;
  var s2cntm = 0;
  var s2i = 0;
  var s2l = sa2.length;
  for (s2i = 0; s2i < s2l; s2i++) {
    if (sa2[s2i].match(/.*[a-z]{2}.*/i)) {
      sa2[s2i] = sa2[s2i].replace(/[^a-z]/i, '');
      s2cnt++;
      for (s1i = 0; s1i < s1cnt; s1i++) {
        if (sa2[s2i] == sa1[s1i]) s2cntm++;
      }
    }
  }
  if (s1cnt < s2cnt) return s2cntm / s1cnt;
  return s2cntm / s2cnt;
}

function getSundayOfMonthOrdinal(date) {
  var dayOfMonth = date.getDate();
  var ordinals = ['First', 'Second', 'Third', 'Fourth', 'Fifth'];
  return ordinals[Math.floor(dayOfMonth / 7)];
}

function getWeekTitles() {

  //get dates for next three Sundays
  var thisSunday = getUpcomingSunday(null, true);
  var nextSunday = dateAdd(thisSunday, 'week', 1);
  var draftSunday = dateAdd(thisSunday, 'week', 2);
  
  //get title for each date
  var thisSundayTitle = fDate(thisSunday, "[ MM.dd ] 'Sunday Announcements'");
  var nextSundayTitle = fDate(nextSunday, "[ MM.dd ] 'Sunday Announcements'");
  var draftSundayTitle = fDate(draftSunday, "[ MM.dd ] 'Sunday Announcements - Draft Document'");
  
  var out = {
    thisSunday  : thisSundayTitle,
    nextSunday  : nextSundayTitle,
    draftSunday : draftSundayTitle,
    dates:{
      thisSunday  : thisSunday,
      nextSunday  : nextSunday,
      draftSunday : draftSunday,
    }
  }

  return out;
}

function escapeGasRegExString(re, escapeCharsArrOpt, ignoreCharsArrOpt){
  
  var charsToReplace = '{[|';

  charsToReplace = charsToReplace.split('');
  if(escapeCharsArrOpt)//add user supplied values
    charsToReplace = charsToReplace.concat(escapeCharsArrOpt);
  if(ignoreCharsArrOpt)//remove user supplied values
    charsToReplace = charsToReplace
    .filter(function(e){return this.toString().indexOf(e)<0;}, ignoreCharsArrOpt);
  
  var str = re.toString()
  .replace(/^\//,'')//remove opening /
  .replace(/\/[imgus]*$/,'');//remove closing / and any flags

  log(charsToReplace)
  for(var c in charsToReplace){
    str = str.replace(charsToReplace[c], '\\$&')
  }
  
  return str;
}

function getDateRegex(dayname, month, date) {
  var regexStr = '';
  var dayRegex = getDayRegexStr(dayname);
  var monthRegex = getMonthRegexStr(month);
  var dateRegex = getDateRegexStr(date);
  //var dateRegex = '(' + date+  ')';
  
  //reference Regex /(Sun|Tues)[- \\/., ]+(Oct|October)[- \\/., ]+(10)/gim
  regexStr = dayRegex + '[- /, ]*' + monthRegex + '[- /, ]*' + dateRegex;
  
  return regexStr;
}

function getDayRegexStr(dayname) {
  switch (dayname) {
    case 'Sunday':
      return '(Sunday|Sunday.|Sun|Sun.| )';
      break;
    case 'Monday':
      return '(Monday|Monday.|Mon|Mon.|Mo|Mo.| )';
      break;
    case 'Tuesday':
      return '(Tuesday|Tuesday.|Tue|Tue.|Tues|Tues.| )';
      break;
    case 'Wednesday':
      return '(Wednesday|Wednesday.|Wed|Wed.| )';
      break;
    case 'Thursday':
      return '(Thursday|Th|Thur|Thu|Thrs|Thrs.|Thurs|Thurs.|Th.|TR|TR.| )';
      break;
    case 'Friday':
      return '(Friday|Friday.|Fri|Fri.|FR|FR.| )';
      break;
    case 'Saturday':
      return '(Saturday|Saturday.|Sat|Sat.| )';
      break;
  }
}

function getMonthRegexStr(month) {
  switch (month) {
    case 0:
      return '(January|January.|Jan|Jan.|1\/)';
      break;
    case 1:
      return '(February|February.|Feb|Feb.|2\/)';
      break;
    case 2:
      return '(March|March.|Mar|Mar.|3\/)';
      break;
    case 3:
      return '(April|April.|Apr|Apr.|4\/)';
      break;
    case 4:
      return '(May|May.|5/)';
      break;
    case 5:
      return '(June|June.|Jun|Jun.|6\/)';
      break;
    case 6:
      return '(July|July.|Jul|Jul.|7\/)';
      break;
    case 7:
      return '(August|August.|Aug|Aug.|8\/)';
      break;
    case 8:
      return '(September|September.|Sep|Sep.|Sept|Sept.|9\/)';
      break;
    case 9:
      return '(October|October.|Oct|Oct.|10/)';
      break;
    case 10:
      return '(November|November.|Nov|Nov.|11\/)';
      break;
    case 11:
      return '(December|December.|Dec|Dec.|12\/)';
      break;
      
  }
}

function getDateRegexStr(date) {
  var dateRegex = '';
  //(19[ ]{1}|(19)+$)
  if (parseInt(date) >= 1 && parseInt(date) <= 9) {
    dateRegex = '(' + date + '[ ]{1}|' + '(' + date + ')+$' + ')';
  } else {
    dateRegex = '(' + date + ')';
  }
  
  return dateRegex;
}

//Takes file name , and returns the date mentioned in it , 
//eg filename 'Copy of [ 10.01 ] Sunday Announcements ' will return Sunday, October 1 2017 00;00:000 
function getDateInName(name) {
  var i1 = name.indexOf('[') + 1;
  var i2 = name.indexOf(']');
  var dateArr = name.substring(i1, i2).trim().split('.');
  // var date = new Date(new Date().getFullYear(),parseInt(dateArr[0])-1,parseInt(dateArr[1]));
  // var month = parseInt(dateArr[0].replace('0',''))-1;
  var month = parseInt(dateArr[0].indexOf('0') == 0 ? dateArr[0].replace('0', '') : dateArr[0]) - 1;
  // var day = parseInt(dateArr[1].replace('0',''));
  var day = parseInt(dateArr[1].indexOf('0') == 0 ? dateArr[1].replace('0', '') : dateArr[1]);
  var date = new Date(new Date().getFullYear(), month, day);
  return date;
}

function extractSundayTitle(doc){
  var search = doc.getBody().findText(config.announcements.sundayPagePattern);
  return search
  ? search.getElement().asText().getText()
  : null;
}
