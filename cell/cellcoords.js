
var colToInt = function(col) {
  var letters = [
    "", "A", "B", "C", "D", "E", "F", "G", 
    "H", "I", "J", "K", "L", "M", "N", "O", "P",
     "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"
  ];
  col = col.trim().split('');
  
  var n = 0;

  for (var i = 0; i < col.length; i++) {
    n *= 26;
    n += letters.indexOf(col[i]);
  }

  return n;
};

var CellCoords = function(cell) {
  cell = cell.split(/([0-9]+)/);
  this.row = parseInt(cell[1]);
  this.column = colToInt(cell[0]);
};

module.exports = CellCoords;
