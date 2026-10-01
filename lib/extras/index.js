/**
 * Formatters and options this fork adds to Carbone: drop, imageSize, fontColor, fontSize, fontBold,
 * fontFamily, pictures from the alternative text, and the options fontOptions, imageUrls,
 * keepLeadingSpace and wrapPrimitiveArrays.
 *
 * The formatters write a marker into the rendered XML and the work is done on the rendered document.
 * Each render has its own random marker, so text coming from the data can never be taken for a
 * marker, and a document that uses none of these formatters is not touched.
 */
var crypto = require('crypto');
var drop = require('./drop');
var font = require('./font');
var image = require('./image');
var xml = require('./xml');

var MARKER_NAMES = ['drop', 'imageSize', 'fontColor', 'fontSize', 'fontBold', 'fontFamily'];

function createState () {
  return {
    marker : crypto.randomBytes(8).toString('hex'),
    used   : {},
    images : {}
  };
}

/**
 * Marker written by a formatter. `context` is the `this` of the formatter (the render options).
 */
function markerOf (context, name) {
  var _state = context !== undefined && context !== null ? context.extras : undefined;
  if (_state === undefined) {
    return ':' + name;
  }
  _state.used[name] = true;
  return ':' + name + _state.marker;
}

function isPlainObject (value) {
  if (value === null || typeof value !== 'object') {
    return false;
  }
  var _prototype = Object.getPrototypeOf(value);
  return _prototype === Object.prototype || _prototype === null;
}

function copyData (value, options) {
  if (typeof value === 'string') {
    return (options.keepLeadingSpace === true && value[0] === ' ') ? '  ' + value.slice(1) : value;
  }
  if (Array.isArray(value) === true) {
    if (options.wrapPrimitiveArrays === true && value.length > 0 && (value[0] === null || typeof value[0] !== 'object')) {
      return value.map(function (item) {
        return { value : copyData(item, options) };
      });
    }
    return value.map(function (item) {
      return copyData(item, options);
    });
  }
  if (isPlainObject(value) === true) {
    var _copy = {};
    Object.keys(value).forEach(function (key) {
      _copy[key] = copyData(value[key], options);
    });
    return _copy;
  }
  return value;
}

/**
 * Data given to the template. The options keepLeadingSpace and wrapPrimitiveArrays work on a copy:
 * the object passed to render is never modified.
 */
function prepareData (data, options) {
  if (options.keepLeadingSpace !== true && options.wrapPrimitiveArrays !== true) {
    return data;
  }
  return copyData(data, options);
}

function removeMarkers (report, state) {
  var _regExp = new RegExp(':(?:' + MARKER_NAMES.join('|') + ')' + state.marker + '\\([^)]*\\)', 'g');
  report.files.forEach(function (file) {
    if (typeof file.data === 'string' && file.data.indexOf(state.marker) !== -1) {
      file.data = file.data.replace(_regExp, '');
    }
  });
}

function applyToDocx (report, options, state, callback) {
  image.replaceAltTextPicturesInDocx(report);
  var _document = xml.findFile(report, 'word/document.xml');
  if (_document === null) {
    return callback(null);
  }
  if (state.used.drop === true) {
    _document.data = drop.dropInDocx(_document.data.toString(), state.marker);
  }
  var _styleRuns = function () {
    if (state.used.fontColor === true || state.used.fontSize === true || state.used.fontBold === true || state.used.fontFamily === true) {
      _document.data = font.styleRunsInDocx(_document.data.toString(), state.marker);
    }
    if (options.fontOptions !== undefined) {
      _document.data = font.styleDocument(_document.data.toString(), options.fontOptions);
    }
    callback(null);
  };
  if (state.used.imageSize === true) {
    return image.insertInDocx(report, state, function (err) {
      if (err) {
        return callback(err);
      }
      _styleRuns();
    });
  }
  _styleRuns();
}

/**
 * Run the work of the extra formatters on the rendered report, before it is zipped.
 */
function apply (report, options, callback) {
  var _state = options.extras;
  var _done = function (err) {
    if (err) {
      return callback(err);
    }
    removeMarkers(report, _state);
    callback(null);
  };
  if (report.extension === 'docx') {
    return applyToDocx(report, options, _state, _done);
  }
  if (report.extension === 'odt') {
    image.replaceAltTextPicturesInOdt(report);
  }
  if (report.extension === 'xlsx' && _state.used.drop === true) {
    report.files.forEach(function (file) {
      if (/^xl\/worksheets\/[^/]+\.xml$/.test(file.name) === true && typeof file.data === 'string') {
        file.data = drop.dropInXlsx(file.data, _state.marker);
      }
    });
  }
  _done(null);
}

module.exports = {
  createState,
  markerOf,
  prepareData,
  apply
};
