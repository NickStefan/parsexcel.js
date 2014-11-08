var ns = require('../cell/namespace');

var na = { 
  value: function() { return ''; },
  text:  function() { return ''; }
};

var Alignments = function(cellNode){
  this.vertical = (cellNode.get('a:alignment/@vertical', ns) || na).value();
  this.horizontal = (cellNode.get('a:alignment/@horizontal', ns) || na).value();
};

module.exports = Alignments;
