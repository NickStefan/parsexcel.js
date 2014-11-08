var ns = require('../cell/namespace');
var getHex = require('./gethex');

var na = { 
  value: function() { return ''; },
  text:  function() { return ''; }
};

var Borders = function(id, borderStyles){
  this.top = (borderStyles[id].get('a:top',ns).attr('style') || na).value();
  this.topColor = getHex((borderStyles[id].get('a:top/a:color/@indexed',ns) || na).value());

  this.bottom = (borderStyles[id].get('a:bottom',ns).attr('style') || na).value();
  this.bottomColor = getHex((borderStyles[id].get('a:bottom/a:color/@indexed',ns) || na).value());  

  this.right = (borderStyles[id].get('a:right',ns).attr('style') || na).value();
  this.rightColor = getHex((borderStyles[id].get('a:right/a:color/@indexed',ns) || na).value());    

  this.left = (borderStyles[id].get('a:left',ns).attr('style') || na).value();
  this.leftColor = getHex((borderStyles[id].get('a:left/a:color/@indexed',ns) || na).value());

  this.diagonal = (borderStyles[id].get('a:diagonal',ns).attr('style') || na).value();
  this.diagonalColor = getHex((borderStyles[id].get('a:diagonal/a:color/@indexed',ns) || na).value());
};

module.exports = Borders;
