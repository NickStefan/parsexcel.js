var Borders = require('./borders');
var Fills = require('./fills');
var Fonts = require('./fonts');
var Alignments = require('./alignments');
var NumFormats = require('./numformats');

var CellStyle = function(cellNode, rawStyles) {
  var fontId = cellNode.attr('fontId').value();
  var fillId = cellNode.attr('fillId').value();
  var borderId = cellNode.attr('borderId').value();
  var numFmtId = cellNode.attr('numFmtId').value();

  this.border = new Borders(borderId, rawStyles.borderStyles);
  this.fill = new Fills(fillId, rawStyles.fillStyles);
  this.font = new Fonts(fontId, rawStyles.fontStyles);
  this.alignment = new Alignments(cellNode);
  this.format = new NumFormats(numFmtId);
};

module.exports = CellStyle;
