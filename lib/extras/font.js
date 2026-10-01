var xml = require('./xml');

// Order of the children of <w:rPr> in the OOXML schema (CT_RPr); Word rejects some other orders.
var RUN_PROPERTY_ORDER = [
  'w:rStyle', 'w:rFonts', 'w:b', 'w:bCs', 'w:i', 'w:iCs', 'w:caps', 'w:smallCaps', 'w:strike', 'w:dstrike',
  'w:outline', 'w:shadow', 'w:emboss', 'w:imprint', 'w:noProof', 'w:snapToGrid', 'w:vanish', 'w:webHidden',
  'w:color', 'w:spacing', 'w:w', 'w:kern', 'w:position', 'w:sz', 'w:szCs', 'w:highlight', 'w:u', 'w:effect',
  'w:bdr', 'w:shd', 'w:fitText', 'w:vertAlign', 'w:rtl', 'w:cs', 'w:em', 'w:lang', 'w:eastAsianLayout',
  'w:specVanish', 'w:oMath'
];

function propertyRank (name) {
  if (name === 'w:rPrChange') {
    return RUN_PROPERTY_ORDER.length + 1;
  }
  var _rank = RUN_PROPERTY_ORDER.indexOf(name);
  return _rank === -1 ? RUN_PROPERTY_ORDER.length : _rank;
}

function parseColor (value) {
  var _color = String(value).trim().replace(/^#/, '');
  if (/^[0-9a-fA-F]{6}$/.test(_color) === true) {
    return _color.toUpperCase();
  }
  return _color === 'auto' ? 'auto' : null;
}

function parseSize (value) {
  var _size = parseFloat(value);
  if (Number.isFinite(_size) === false || _size <= 0 || _size > 1638) {
    return null;
  }
  return Math.round(_size * 2);
}

function parseFamily (value) {
  var _family = String(value).trim();
  var _hasControlCharacter = _family.split('').some(function (character) {
    return character.charCodeAt(0) < 32;
  });
  if (_family.length === 0 || _family.length > 64 || /["<>&]/.test(_family) === true || _hasControlCharacter === true) {
    return null;
  }
  return _family;
}

/**
 * New <w:rPr> children for the styles of one run, keyed by the element names they replace.
 */
function runProperties (styles) {
  var _properties = {};
  styles.forEach(function (style) {
    if (style.name === 'fontColor') {
      var _color = parseColor(style.value);
      if (_color !== null) {
        _properties['w:color'] = '<w:color w:val="' + _color + '"/>';
      }
    }
    else if (style.name === 'fontSize') {
      var _size = parseSize(style.value);
      if (_size !== null) {
        _properties['w:sz'] = '<w:sz w:val="' + _size + '"/>';
        _properties['w:szCs'] = '<w:szCs w:val="' + _size + '"/>';
      }
    }
    else if (style.name === 'fontBold') {
      _properties['w:b'] = '<w:b/>';
      _properties['w:bCs'] = '<w:bCs/>';
    }
    else if (style.name === 'fontFamily') {
      var _family = parseFamily(style.value);
      if (_family !== null) {
        _properties['w:rFonts'] = '<w:rFonts w:ascii="' + _family + '" w:hAnsi="' + _family + '" w:eastAsia="' + _family + '" w:cs="' + _family + '"/>';
      }
    }
  });
  return _properties;
}

function styleRun (runXml, properties) {
  if (Object.keys(properties).length === 0) {
    return runXml;
  }
  var _children = xml.childElements(runXml, 0, runXml.length);
  var _rPr = _children.find(function (child) {
    return child.name === 'w:rPr';
  });
  var _existing = [];
  if (_rPr !== undefined && runXml.slice(_rPr.start, _rPr.end).endsWith('/>') === false) {
    _existing = xml.childElements(runXml, _rPr.start, _rPr.end).map(function (child) {
      return { name : child.name, text : runXml.slice(child.start, child.end) };
    });
  }
  var _merged = _existing.filter(function (child) {
    return properties[child.name] === undefined;
  });
  Object.keys(properties).forEach(function (name) {
    _merged.push({ name : name, text : properties[name] });
  });
  _merged = _merged.map(function (child, index) {
    return { child : child, index : index };
  }).sort(function (a, b) {
    return (propertyRank(a.child.name) - propertyRank(b.child.name)) || (a.index - b.index);
  }).map(function (item) {
    return item.child.text;
  });
  var _newRPr = '<w:rPr>' + _merged.join('') + '</w:rPr>';
  if (_rPr !== undefined) {
    return runXml.slice(0, _rPr.start) + _newRPr + runXml.slice(_rPr.end);
  }
  var _openingEnd = runXml.indexOf('>') + 1;
  return runXml.slice(0, _openingEnd) + _newRPr + runXml.slice(_openingEnd);
}

/**
 * Apply the styles written by the font formatters to the run that holds them and remove the markers.
 */
function styleRunsInDocx (documentXml, marker) {
  var _markerSource = ':(fontColor|fontSize|fontBold|fontFamily)' + marker + '\\(([^)]*)\\)';
  var _first = new RegExp(_markerSource);
  var _match = _first.exec(documentXml);
  for (; _match !== null; _match = _first.exec(documentXml)) {
    var _run = xml.findEnclosing(documentXml, _match.index, 'w:r');
    if (_run === null) {
      documentXml = documentXml.slice(0, _match.index) + documentXml.slice(_match.index + _match[0].length);
      continue;
    }
    var _runXml = documentXml.slice(_run.start, _run.end);
    var _styles = [];
    var _cleanRun = _runXml.replace(new RegExp(_markerSource, 'g'), function (whole, name, value) {
      _styles.push({ name : name, value : value });
      return '';
    });
    documentXml = documentXml.slice(0, _run.start) + styleRun(_cleanRun, runProperties(_styles)) + documentXml.slice(_run.end);
  }
  return documentXml;
}

/**
 * Option `fontOptions`: change the font family, size or weight written in every run of the document.
 */
function styleDocument (documentXml, fontOptions) {
  if (fontOptions === undefined || fontOptions === null) {
    return documentXml;
  }
  if (fontOptions.fontFamily !== undefined) {
    var _family = parseFamily(fontOptions.fontFamily);
    if (_family !== null) {
      documentXml = documentXml.replace(/<w:rFonts\s[^>]*>/g, function (element) {
        return element.replace(/(w:(?:ascii|hAnsi|eastAsia|cs)=")[^"]*(")/g, function (whole, before, after) {
          return before + _family + after;
        });
      });
    }
  }
  if (fontOptions.fontSize !== undefined) {
    var _size = parseSize(fontOptions.fontSize);
    if (_size !== null) {
      documentXml = documentXml.replace(/(<w:(?:sz|szCs) w:val=")[^"]*(")/g, '$1' + _size + '$2');
    }
  }
  if (fontOptions.fontBold !== undefined) {
    var _bold = (fontOptions.fontBold === true || fontOptions.fontBold === 'true') ? 'true' : 'false';
    documentXml = documentXml.replace(/(<w:(?:b|bCs) w:val=")[^"]*(")/g, '$1' + _bold + '$2');
  }
  return documentXml;
}

module.exports = {
  styleRunsInDocx,
  styleDocument
};
