var Borders = require('./borders');
var Fills = require('./fills');
var Fonts = require('./fonts');
var Alignments = require('./alignments');

var CellStyle = function(cellNode, rawStyles) {
  var fontId = cellNode.attr('fontId').value(),
    fillId = cellNode.attr('fillId').value(),
    borderId = cellNode.attr('borderId').value();

  this.border = new Borders(borderId, rawStyles.borderStyles);
  this.fill = new Fills(fillId, rawStyles.fillStyles);
  this.font = new Fonts(fontId, rawStyles.fontStyles);
  this.alignment = new Alignments(cellNode);
};

module.exports = CellStyle;
