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

function fDate(date, format){//returns the date formatted with format, default to today if date not provided
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