var na = { 
  value: function() { return ''; },
  text:  function() { return ''; }
};

var ns = {a: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'};

var Alignments = function(cellNode){
  this.vertical = (cellNode.get('a:alignment/@vertical', ns) || na).value();
  this.horizontal = (cellNode.get('a:alignment/@horizontal', ns) || na).value();
};

module.exports = Alignments;
