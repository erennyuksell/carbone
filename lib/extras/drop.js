var xml = require('./xml');

var DOCX_TAGS = {
  p   : 'w:p',
  tr  : 'w:tr',
  tbl : 'w:tbl'
};

function markerRegExp (marker) {
  return new RegExp(':drop' + marker + '\\((\\w*)\\)');
}

/**
 * Remove every element marked by the `drop` formatter in a Word document.
 * Each pass removes at least the marker itself, so the loop always ends.
 */
function dropInDocx (documentXml, marker) {
  var _regExp = markerRegExp(marker);
  var _match = _regExp.exec(documentXml);
  for (; _match !== null; _match = _regExp.exec(documentXml)) {
    var _tag = _match[1];
    var _result = null;
    if (_tag === 'tc') {
      _result = dropColumn(documentXml, _match.index);
    }
    else if (DOCX_TAGS[_tag] !== undefined) {
      var _element = xml.findEnclosing(documentXml, _match.index, DOCX_TAGS[_tag]);
      if (_element !== null) {
        _result = documentXml.slice(0, _element.start) + documentXml.slice(_element.end);
      }
    }
    if (_result === null) {
      _result = documentXml.slice(0, _match.index) + documentXml.slice(_match.index + _match[0].length);
    }
    documentXml = _result;
  }
  return documentXml;
}

function gridSpan (documentXml, tcElement) {
  var _cell = documentXml.slice(tcElement.start, tcElement.end);
  var _properties = /^<w:tc(?:\s[^>]*)?>\s*<w:tcPr>([\s\S]*?)<\/w:tcPr>/.exec(_cell);
  var _span = _properties ? /<w:gridSpan w:val="(\d+)"/.exec(_properties[1]) : null;
  return _span ? parseInt(_span[1], 10) : 1;
}

/**
 * Remove the table column that holds the marker from every row of its table, then share the
 * width of the removed column between the remaining ones. A cell spanning the removed column
 * loses one span instead of being removed.
 */
function dropColumn (documentXml, position) {
  var _cell = xml.findEnclosing(documentXml, position, 'w:tc');
  var _row = xml.findEnclosing(documentXml, position, 'w:tr');
  var _table = xml.findEnclosing(documentXml, position, 'w:tbl');
  if (_cell === null || _row === null || _table === null) {
    return null;
  }
  var _column = 0;
  var _cells = xml.outerElements(documentXml, 'w:tc', _row.start, _row.end);
  for (var c = 0; c < _cells.length && _cells[c].start !== _cell.start; c++) {
    _column += gridSpan(documentXml, _cells[c]);
  }
  var _edits = [];
  var _rows = xml.outerElements(documentXml, 'w:tr', _table.start, _table.end);
  _rows.forEach(function (row) {
    var _firstColumn = 0;
    var _rowCells = xml.outerElements(documentXml, 'w:tc', row.start, row.end);
    for (var i = 0; i < _rowCells.length; i++) {
      var _span = gridSpan(documentXml, _rowCells[i]);
      if (_column >= _firstColumn && _column < _firstColumn + _span) {
        if (_span > 1) {
          var _cellXml = documentXml.slice(_rowCells[i].start, _rowCells[i].end);
          _edits.push({
            start : _rowCells[i].start,
            end   : _rowCells[i].end,
            text  : _cellXml.replace('<w:gridSpan w:val="' + _span + '"', '<w:gridSpan w:val="' + (_span - 1) + '"')
          });
        }
        else {
          _edits.push({ start : _rowCells[i].start, end : _rowCells[i].end, text : '' });
        }
        break;
      }
      _firstColumn += _span;
    }
  });
  var _grid = xml.outerElements(documentXml, 'w:tblGrid', _table.start, _table.end)[0];
  if (_grid !== undefined) {
    _edits.push({ start : _grid.start, end : _grid.end, text : resizeGrid(documentXml.slice(_grid.start, _grid.end), _column) });
  }
  return xml.applyEdits(documentXml, _edits);
}

function resizeGrid (gridXml, removedColumn) {
  var _columns = xml.outerElements(gridXml, 'w:gridCol', 0, gridXml.length);
  if (removedColumn >= _columns.length) {
    return gridXml;
  }
  var _widthRegExp = /w:w="(\d+)"/;
  var _widths = _columns.map(function (column) {
    var _width = _widthRegExp.exec(gridXml.slice(column.start, column.end));
    return _width ? parseInt(_width[1], 10) : 0;
  });
  var _total = _widths.reduce(function (sum, width) {
    return sum + width;
  }, 0);
  var _remaining = _total - _widths[removedColumn];
  var _edits = [{ start : _columns[removedColumn].start, end : _columns[removedColumn].end, text : '' }];
  if (_remaining > 0) {
    _columns.forEach(function (column, index) {
      if (index === removedColumn) {
        return;
      }
      var _width = Math.floor(_widths[index] * _total / _remaining);
      _edits.push({
        start : column.start,
        end   : column.end,
        text  : gridXml.slice(column.start, column.end).replace(_widthRegExp, 'w:w="' + _width + '"')
      });
    });
  }
  return xml.applyEdits(gridXml, _edits);
}

/**
 * Remove every row marked by `drop(row)` in a worksheet.
 */
function dropInXlsx (sheetXml, marker) {
  var _regExp = markerRegExp(marker);
  var _match = _regExp.exec(sheetXml);
  for (; _match !== null; _match = _regExp.exec(sheetXml)) {
    var _element = _match[1] === 'row' ? xml.findEnclosing(sheetXml, _match.index, 'row') : null;
    if (_element !== null) {
      sheetXml = sheetXml.slice(0, _element.start) + sheetXml.slice(_element.end);
    }
    else {
      sheetXml = sheetXml.slice(0, _match.index) + sheetXml.slice(_match.index + _match[0].length);
    }
  }
  return sheetXml;
}

module.exports = {
  dropInDocx,
  dropInXlsx
};
