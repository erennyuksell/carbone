/**
 * String helpers for the fork's extra formatters. They work on the XML text directly, without
 * parsing and serializing it again, so everything they do not touch stays byte for byte the same.
 */

function escapeRegExp (text) {
  return text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/**
 * Regular expression matching opening, closing and self-closing tags of one element name.
 * The lookahead keeps `<w:p` from matching `<w:pPr>` or `<w:pStyle>`.
 */
function tagRegExp (tagName) {
  return new RegExp('<(/?)' + escapeRegExp(tagName) + '(?=[\\s/>])[^>]*>', 'g');
}

/**
 * Innermost element `tagName` that contains `position`.
 *
 * @return {Object|null} { start, end } where xml.slice(start, end) is the whole element
 */
function findEnclosing (xml, position, tagName) {
  var _regExp = tagRegExp(tagName);
  var _openPositions = [];
  var _target = -1;
  var _targetDepth = -1;
  var _match = _regExp.exec(xml);
  for (; _match !== null; _match = _regExp.exec(xml)) {
    if (_target === -1 && _match.index > position) {
      if (_openPositions.length === 0) {
        return null;
      }
      _target = _openPositions[_openPositions.length - 1];
      _targetDepth = _openPositions.length;
    }
    if (_match[1] === '/') {
      if (_target !== -1 && _openPositions.length === _targetDepth) {
        return { start : _target, end : _match.index + _match[0].length };
      }
      _openPositions.pop();
    }
    else if (_match[0].endsWith('/>') === false) {
      _openPositions.push(_match.index);
    }
  }
  return null;
}

/**
 * Outermost elements `tagName` between `from` and `to`. Elements nested in another element of
 * the same name are skipped: the rows of a nested table are not rows of the outer table.
 *
 * @return {Array} [{ start, end }]
 */
function outerElements (xml, tagName, from, to) {
  var _regExp = tagRegExp(tagName);
  var _elements = [];
  var _depth = 0;
  var _start = -1;
  _regExp.lastIndex = from;
  var _match = _regExp.exec(xml);
  for (; _match !== null && _match.index < to; _match = _regExp.exec(xml)) {
    var _end = _match.index + _match[0].length;
    if (_match[1] === '/') {
      _depth--;
      if (_depth === 0) {
        _elements.push({ start : _start, end : _end });
      }
    }
    else if (_match[0].endsWith('/>') === true) {
      if (_depth === 0) {
        _elements.push({ start : _match.index, end : _end });
      }
    }
    else {
      if (_depth === 0) {
        _start = _match.index;
      }
      _depth++;
    }
  }
  return _elements;
}

/**
 * Direct children of the element that starts at `start` and ends at `end`, whatever their name.
 *
 * @return {Array} [{ name, start, end }]
 */
function childElements (xml, start, end) {
  var _regExp = /<(\/?)([\w.:-]+)[^>]*>/g;
  var _children = [];
  var _depth = 0;
  var _childStart = -1;
  var _childName = '';
  _regExp.lastIndex = xml.indexOf('>', start) + 1;
  var _match = _regExp.exec(xml);
  for (; _match !== null && _match.index < end; _match = _regExp.exec(xml)) {
    var _tagEnd = _match.index + _match[0].length;
    if (_match[1] === '/') {
      _depth--;
      if (_depth === 0) {
        _children.push({ name : _childName, start : _childStart, end : _tagEnd });
      }
      else if (_depth < 0) {
        break;
      }
    }
    else if (_match[0].endsWith('/>') === true) {
      if (_depth === 0) {
        _children.push({ name : _match[2], start : _match.index, end : _tagEnd });
      }
    }
    else {
      if (_depth === 0) {
        _childStart = _match.index;
        _childName = _match[2];
      }
      _depth++;
    }
  }
  return _children;
}

/**
 * Apply edits { start, end, text } to xml. Edits must not overlap.
 */
function applyEdits (xml, edits) {
  var _sorted = edits.slice().sort(function (a, b) {
    return b.start - a.start;
  });
  for (var i = 0; i < _sorted.length; i++) {
    xml = xml.slice(0, _sorted[i].start) + _sorted[i].text + xml.slice(_sorted[i].end);
  }
  return xml;
}

function findFile (report, name) {
  for (var i = 0; i < report.files.length; i++) {
    if (report.files[i].name === name) {
      return report.files[i];
    }
  }
  return null;
}

/**
 * Add `<Default Extension>` to [Content_Types].xml unless the extension is already declared:
 * Word refuses a package that declares the same extension twice.
 */
function ensureContentType (report, extension, contentType) {
  var _file = findFile(report, '[Content_Types].xml');
  if (_file === null) {
    return;
  }
  var _xml = _file.data.toString();
  if (new RegExp('<Default [^>]*Extension="' + escapeRegExp(extension) + '"', 'i').test(_xml) === true) {
    return;
  }
  _file.data = _xml.replace('</Types>', '<Default Extension="' + extension + '" ContentType="' + contentType + '"/></Types>');
}

/**
 * Add an image relationship to word/_rels/document.xml.rels and return its id.
 */
function addImageRelationship (report, target) {
  var _file = findFile(report, 'word/_rels/document.xml.rels');
  if (_file === null) {
    return null;
  }
  var _xml = _file.data.toString();
  var _index = 1;
  while (_xml.indexOf('Id="rIdCarboneImg' + _index + '"') !== -1) {
    _index++;
  }
  var _id = 'rIdCarboneImg' + _index;
  _file.data = _xml.replace('</Relationships>',
    '<Relationship Id="' + _id + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="' + target + '"/></Relationships>');
  return _id;
}

module.exports = {
  escapeRegExp,
  findEnclosing,
  outerElements,
  childElements,
  applyEdits,
  findFile,
  ensureContentType,
  addImageRelationship
};
