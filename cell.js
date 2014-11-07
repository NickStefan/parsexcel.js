var CellCoords = require('./cellcoords');

var ns = {a: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'};
var na = { 
  value: function() { return ''; },
  text:  function() { return ''; }
};

var Cell = function(cellNode) {
  var r = cellNode.attr('r').value();
  var type = (cellNode.attr('t') || na).value();
  var value = (cellNode.get('a:v', ns) || na ).text();
  var formula = (cellNode.get('a:f', ns) || na).text();
  var style = (cellNode.attr('s') || na).value();
  var coords = new CellCoords(r);

  this.column = coords.column;
  this.row = coords.row;
  this.value = value;
  this.formula = formula;
  this.type = type;
  this.style = style;
};

module.exports = Cell;
