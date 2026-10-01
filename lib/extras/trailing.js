/**
 * Option trimTrailingParagraphs: remove the empty paragraphs at the end of a Word document.
 *
 * A document often ends with empty paragraphs: blank lines typed after the last table, or paragraphs that only held
 * a marker. When the content fills the page, they are pushed to a new page and the PDF ends with a blank page.
 *
 * Only paragraphs that print nothing are removed: no text other than spaces, no picture, field, symbol, bookmark,
 * comment or section break. If the document then ends with a table, one paragraph of 1 point stays after it, because
 * Word requires a paragraph after a table at the end of the body.
 */

/* What makes a paragraph print something, or carry something that must not be lost */
var KEEP = /<(?:w:drawing|w:pict|w:object|mc:AlternateContent|w:fldChar|w:fldSimple|w:instrText|w:sym|w:noBreakHyphen|w:softHyphen|w:footnoteReference|w:endnoteReference|w:commentReference|w:commentRangeStart|w:commentRangeEnd|w:bookmarkStart|w:bookmarkEnd|w:sectPr|w:ptab|v:shape|w:ruby)\b/;
var TEXT = /<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/g;
var BLANK = /^[\s  -​　]*$/;
var TAG = /<(\/?)([\w:.-]+)(?:\s[^>]*?)?(\/?)>/g;

/* The paragraph Word needs after a table at the end of the body, as small as possible */
var MINIMAL_PARAGRAPH = '<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="20" w:lineRule="exact"/>'
  + '<w:rPr><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr></w:pPr></w:p>';

function decode (text) {
  return text.replace(/&#(\d+);/g, function (match, code) {
    return String.fromCharCode(parseInt(code, 10));
  }).replace(/&#x([0-9a-f]+);/gi, function (match, code) {
    return String.fromCharCode(parseInt(code, 16));
  }).replace(/&nbsp;/g, ' ');
}

/**
 * @param  {String}  paragraph XML of a w:p element
 * @return {Boolean}           true if the paragraph prints nothing and carries nothing to keep
 */
function isEmptyParagraph (paragraph) {
  if (KEEP.test(paragraph) === true) {
    return false;
  }
  var _match;
  TEXT.lastIndex = 0;
  while ((_match = TEXT.exec(paragraph)) !== null) {
    if (BLANK.test(decode(_match[1])) === false) {
      return false;
    }
  }
  return true;
}

/**
 * Direct children of the body, in order: { name, start, end } positions in bodyXml
 */
function bodyChildren (bodyXml) {
  var _children = [];
  var _depth = 0;
  var _start = -1;
  var _name = '';
  var _match;
  TAG.lastIndex = 0;
  while ((_match = TAG.exec(bodyXml)) !== null) {
    var _isClosing = _match[1] === '/';
    var _isSelfClosing = _match[3] === '/';
    if (_isClosing === false && _depth === 0) {
      _start = _match.index;
      _name = _match[2];
    }
    if (_isSelfClosing === true) {
      if (_depth === 0) {
        _children.push({ name : _name, start : _start, end : TAG.lastIndex });
      }
    }
    else if (_isClosing === true) {
      _depth--;
      if (_depth === 0) {
        _children.push({ name : _name, start : _start, end : TAG.lastIndex });
      }
    }
    else {
      _depth++;
    }
  }
  return _children;
}

/**
 * @param  {String} documentXml word/document.xml
 * @return {String}             the same document without its trailing empty paragraphs
 */
function trimTrailingParagraphs (documentXml) {
  var _bodyStart = documentXml.indexOf('<w:body>');
  var _bodyEnd = documentXml.lastIndexOf('</w:body>');
  if (_bodyStart === -1 || _bodyEnd === -1) {
    return documentXml;
  }
  _bodyStart += '<w:body>'.length;
  var _body = documentXml.slice(_bodyStart, _bodyEnd);
  var _children = bodyChildren(_body);
  // the section properties of the last section end the body
  var _last = _children.length - 1;
  if (_last >= 0 && _children[_last].name === 'w:sectPr') {
    _last--;
  }
  var _removeFrom = -1;
  // never remove the first element: a document keeps at least one paragraph
  for (var i = _last; i > 0; i--) {
    var _child = _children[i];
    if (_child.name !== 'w:p' || isEmptyParagraph(_body.slice(_child.start, _child.end)) === false) {
      break;
    }
    _removeFrom = i;
  }
  if (_removeFrom === -1) {
    return documentXml;
  }
  var _before = _children[_removeFrom - 1];
  var _replacement = _before.name === 'w:tbl' ? MINIMAL_PARAGRAPH : '';
  var _trimmedBody = _body.slice(0, _children[_removeFrom].start) + _replacement + _body.slice(_children[_last].end);
  return documentXml.slice(0, _bodyStart) + _trimmedBody + documentXml.slice(_bodyEnd);
}

module.exports = {
  trimTrailingParagraphs,
  isEmptyParagraph
};
