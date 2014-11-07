var _ = require('underscore');
var CellCoords = require('./cellcoords');

var calculateDimensions = function(cells) {
  var comparator = function (a, b) { return a-b; };
  var allRows = _(cells).map(function (cell) { return cell.row; }).sort(comparator);
  var allCols = _(cells).map(function (cell) { return cell.column; }).sort(comparator);
  var minRow = allRows[0];
  var maxRow = _.last(allRows);
  var minCol = allCols[0];
  var maxCol = _.last(allCols);

  return [
    {row: minRow, column: minCol},
    {row: maxRow, column: maxCol}
  ];
};

var ns = {a: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'};

var calcSheetDimensions = function(sheet){
  var d = sheet.data.get('//a:dimension/@ref', ns);
  if (d) {
    sheet.dimensions = _.map(d.value().split(':'), function(v) { return new CellCoords(v); });
  } else {
    sheet.dimensions = calculateDimensions(sheet.cells)
  }
}

module.exports = calcSheetDimensions;
