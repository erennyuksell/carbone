var dayjs = require('dayjs');
// the plugins these formatters use, in the order of lib/index.js (dayjs installs a plugin once)
dayjs.extend(require('dayjs/plugin/advancedFormat'));
dayjs.extend(require('dayjs/plugin/localizedFormat'));
dayjs.extend(require('dayjs/plugin/customParseFormat'));
dayjs.extend(require('dayjs/plugin/utc'));
dayjs.extend(require('dayjs/plugin/isoWeek'));
dayjs.extend(require('dayjs/plugin/timezone'));


/**
 * Format dates. It takes an output date pattern as an argument. Date patterns are available on [this section](#date-formats).
 * It is possible to change the timezone through the option `options.timezone` and the lang through `options.lang`.
 * List of timezones: [https://en.wikipedia.org/wiki/List_of_tz_database_time_zones](https://en.wikipedia.org/wiki/List_of_tz_database_time_zones).
 *
 * @version 3.0.0 updated
 *
 * @exampleContext {"lang":"en", "timezone":"Europe/Paris"}
 * @example ["20160131", "L"]
 * @example ["20160131", "LL"]
 * @example ["20160131", "LLLL"]
 * @example ["20160131", "dddd"]
 *
 * @exampleContext {"lang":"fr", "timezone":"Europe/Paris"}
 * @example ["2017-05-10T15:57:23.769561+03:00", "LLLL"]
 * @example ["2017-05-10 15:57:23.769561+03:00", "LLLL"]
 * @example ["20160131", "LLLL"]
 * @example ["20160131", "dddd"]
 *
 * @exampleContext {"lang":"fr", "timezone":"Europe/Paris"}
 * @example ["20160131", "dddd", "YYYYMMDD"]
 * @example [1410715640, "LLLL", "X" ]
 *
 * @exampleContext {"lang":"fr", "timezone": "Asia/Singapore"}
 * @example ["20160131", "dddd", "YYYYMMDD"]
 * @example [1410715640, "LLLL", "X" ]
 *
 * @param  {String|Number} d   date to format
 * @param  {String} patternOut output format
 * @param  {String} patternIn  [optional] input format, "ISO 8601" by default
 * @return {String}            return formatted date
 */
function formatD (d, patternOut, patternIn) {
  if (d !== null && typeof d !== 'undefined') {
    var _date = parse(d, patternIn).tz(this.timezone).locale(this.lang);
    return _date.format(isoWeekAsText(_date, patternOut));
  }
  return d;
}

/**
 *
 * Add a time to a date. Available units: day, week,	month, quarter, year, hour, minute, second and millisecond.
 * Units are case insensitive, and support plural and short forms.
 *
 * @version 3.0.0 new
 *
 * @exampleContext {"lang":"fr", "timezone":"Europe/Paris"}
 * @example ["2017-05-10T15:57:23.769561+03:00", "3", "day"]
 * @example ["2017-05-10 15:57:23.769561+03:00", "3", "month"]
 * @example ["20160131", "3", "day"]
 * @example ["20160131", "3", "month"]
 * @example ["31-2016-01", "3", "month", "DD-YYYY-MM"]
 *
 * @param      {String|Number}  d   input date
 * @param      {Number}  amount     The amount
 * @param      {String}  unit       The unit
 * @param      {String}  patternIn  [optional] input format, ISO8601 by default
 * @return     {Date}               return a date, which can be formatted with formatD, or manipulated with other formatters
 */
function addD (d, amount, unit, patternIn) {
  if (d !== null && typeof d !== 'undefined') {
    return onWallTime(parse(d, patternIn), function (date) {
      return date.add(parseInt(amount, 10), unit || 'day');
    });
  }
  return d;
}

/**
 *
 * Subtract a time to a date. Available units: day, week,	month, quarter, year, hour, minute, second and millisecond.
 * Units are case insensitive, and support plural and short forms.
 *
 * @version 3.0.0 new
 *
 * @exampleContext {"lang":"fr", "timezone":"Europe/Paris"}
 * @example ["2017-05-10T15:57:23.769561+03:00", "3", "day"]
 * @example ["2017-05-10 15:57:23.769561+03:00", "3", "month"]
 * @example ["20160131", "3", "day"]
 * @example ["20160131", "3", "month"]
 * @example ["31-2016-01", "3", "month", "DD-YYYY-MM"]
 *
 * @param      {String|Number}  d   input date
 * @param      {Number}  amount     The amount
 * @param      {String}  unit       The unit
 * @param      {String}  patternIn  [optional] input format, ISO8601 by default
 * @return     {Date}               return a date, which can be formatted with formatD, or manipulated with other formatters
 */
function subD (d, amount, unit, patternIn) {
  if (d !== null && typeof d !== 'undefined') {
    return onWallTime(parse(d, patternIn), function (date) {
      return date.subtract(parseInt(amount, 10), unit || 'day');
    });
  }
  return d;
}

/**
 *
 * Create a date and set it to the start of a unit of time.
 *
 * @version 3.0.0 new
 *
 * @exampleContext {"lang":"fr", "timezone":"Europe/Paris"}
 * @example ["2017-05-10T15:57:23.769561+03:00", "day"]
 * @example ["2017-05-10 15:57:23.769561+03:00", "month"]
 * @example ["20160131", "day"]
 * @example ["20160131", "month"]
 * @example ["31-2016-01", "month", "DD-YYYY-MM"]
 *
 * @param      {String|Number}  d   input date
 * @param      {String}  unit       The unit
 * @param      {String}  patternIn  [optional] input format, ISO8601 by default
 * @return     {Date}               return a date, which can be formatted with formatD, or manipulated with other formatters
 */
function startOfD (d, unit, patternIn) {
  if (d !== null && typeof d !== 'undefined') {
    return onWallTime(parse(d, patternIn), function (date) {
      return date.startOf(unit || 'year');
    });
  }
  return d;
}

/**
 *
 * Create a date and set it to the end of a unit of time.
 *
 * @version 3.0.0 new
 *
 * @exampleContext {"lang":"fr", "timezone":"Europe/Paris"}
 * @example ["2017-05-10T15:57:23.769561+03:00", "day"]
 * @example ["2017-05-10 15:57:23.769561+03:00", "month"]
 * @example ["20160131", "day"]
 * @example ["20160131", "month"]
 * @example ["31-2016-01", "month", "DD-YYYY-MM"]
 *
 * @param      {String|Number}  d   input date
 * @param      {String}  unit       The unit
 * @param      {String}  patternIn  [optional] input format, ISO8601 by default
 * @return     {Date}               return a date, which can be formatted with formatD, or manipulated with other formatters
 */
function endOfD (d, unit, patternIn) {
  if (d !== null && typeof d !== 'undefined') {
    return onWallTime(parse(d, patternIn), function (date) {
      return date.endOf(unit || 'year');
    });
  }
  return d;
}


/**
 * Format dates
 *
 * @deprecated
 * @version 1.0.0 deprecated
 *
 * @exampleContext {"lang":"en", "timezone":"Europe/Paris"}
 * @example ["20160131", "YYYYMMDD", "L"]
 * @example ["20160131", "YYYYMMDD", "LL"]
 * @example ["20160131", "YYYYMMDD", "LLLL"]
 * @example ["20160131", "YYYYMMDD", "dddd"]
 * @example [1410715640, "X", "LLLL"]
 *
 * @exampleContext {"lang":"fr", "timezone":"Europe/Paris"}
 * @example ["20160131", "YYYYMMDD", "LLLL"]
 * @example ["20160131", "YYYYMMDD", "dddd"]
 *
 * @param  {String|Number} d   date to format
 * @param  {String} patternIn  input format
 * @param  {String} patternOut output format
 * @return {String}            return formatted date
 */
function convDate (d, patternIn, patternOut) {
  return formatD.call(this, d, patternOut, patternIn);
}


/**
 * Timezone of a date written without offset ("20160131", "2016-01-31 10:00"): Carbone reads it as a date and time of
 * Europe/Paris, then converts it to the output timezone (options.timezone). It used to come from the timezone plugin
 * of dayjs; with dayjs/plugin/timezone, dayjs(text) reads the text in the timezone of the server, and the same
 * template printed another day on a server in UTC or in Istanbul.
 */
var INPUT_TIMEZONE = 'Europe/Paris';
/* A time followed by an offset: "15:57:23Z", "15:57:23.769+03:00", "15:57 -0800" (not the year of "06-01-2014") */
var TIME_WITH_OFFSET = /\d{2}:\d{2}(?::\d{2}(?:[.,]\d+)?)?\s*(?:Z|[+-]\d{2}(?::?\d{2})?)$/i;

/**
 * Read a date. A text with a time and an offset ("...T10:00:00Z", "...10:00+03:00"), a text read with an offset token
 * (Z, ZZ), a number and a unix timestamp ("X", "x") are instants; any other text is read in INPUT_TIMEZONE.
 *
 * @private
 * @param      {string}  d          not undefined/null date
 * @param      {string}  patternIn  The pattern
 * @return     {Object}             dayjs
 */
function parse (d, patternIn) {
  // if the date is already parsed
  if (typeof(d) === 'object' && d.isValid) {
    return d;
  }
  var _text = d + '';
  var _isInstant = typeof(d) === 'number' || patternIn === 'X' || patternIn === 'x' || /Z/.test(patternIn || '') || TIME_WITH_OFFSET.test(_text.trim());
  if (_isInstant === true) {
    return patternIn ? dayjs(_text, patternIn) : dayjs(_text);
  }
  // read the fields of the text as they are (dayjs.tz(text, pattern, zone) reads them in the server's timezone),
  // then place this date and time in INPUT_TIMEZONE
  var _wallTime = patternIn ? dayjs.utc(_text, patternIn) : dayjs.utc(_text);
  if (_wallTime.isValid() === false) {
    // an invalid text stays invalid, as with dayjs(text)
    return patternIn ? dayjs(_text, patternIn) : dayjs(_text);
  }
  return dayjs.tz(_wallTime.format('YYYY-MM-DDTHH:mm:ss.SSS'), INPUT_TIMEZONE);
}



/**
 * Calendar operation (add days, start of month...) on the date and time of INPUT_TIMEZONE, then read back in that
 * timezone. Done on the server's own clock, it gave another hour or day when the server is not in INPUT_TIMEZONE, and
 * around a change of summer time.
 *
 * @private
 * @param  {Object}   date      dayjs
 * @param  {Function} operation function(dayjs in UTC holding the wall time) => dayjs
 * @return {Object}             dayjs
 */
function onWallTime (date, operation) {
  if (date.isValid() === false) {
    return date;
  }
  var _wallTime = date.tz(INPUT_TIMEZONE).format('YYYY-MM-DDTHH:mm:ss.SSS');
  var _result = operation(dayjs.utc(_wallTime));
  return dayjs.tz(_result.format('YYYY-MM-DDTHH:mm:ss.SSS'), INPUT_TIMEZONE);
}

/**
 * dayjs computes the ISO week (W, WW, GGGG) of a date in a timezone with the server's clock: on a server in UTC, a
 * Wednesday at midnight in Paris was in the previous week. The tokens are replaced by the week of the date and time
 * of the timezone, as text.
 *
 * @private
 * @param  {Object} date    dayjs in the output timezone
 * @param  {String} pattern output pattern
 * @return {String}         pattern without ISO week tokens
 */
function isoWeekAsText (date, pattern) {
  if (typeof(pattern) !== 'string' || /W|GGGG/.test(pattern) === false) {
    return pattern;
  }
  var _wallTime = dayjs.utc(date.format('YYYY-MM-DDTHH:mm:ss.SSS'));
  return pattern.replace(/\[[^\]]*]|WW|W|GGGG/g, function (token) {
    if (token[0] === '[') {
      return token;
    }
    if (token === 'GGGG') {
      return '[' + _wallTime.isoWeekYear() + ']';
    }
    var _week = String(_wallTime.isoWeek());
    return '[' + (token === 'WW' && _week.length < 2 ? '0' + _week : _week) + ']';
  });
}

module.exports = {
  formatD,
  convDate,
  addD,
  subD,
  startOfD,
  endOfD
};
