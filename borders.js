var na = { 
  value: function() { return ''; },
  text:  function() { return ''; }
};

var ns = {a: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'};

var Borders = function(id, borderStyles){
  this.top = (borderStyles[id].get('a:top',ns).attr('style') || na).value();
  this.topColor = (borderStyles[id].get('a:top/a:color/@indexed',ns) || na).value();

  this.bottom = (borderStyles[id].get('a:bottom',ns).attr('style') || na).value();
  this.bottomColor = (borderStyles[id].get('a:bottom/a:color/@indexed',ns) || na).value();  

  this.right = (borderStyles[id].get('a:right',ns).attr('style') || na).value();
  this.rightColor = (borderStyles[id].get('a:right/a:color/@indexed',ns) || na).value();    

  this.left = (borderStyles[id].get('a:left',ns).attr('style') || na).value();
  this.leftColor = (borderStyles[id].get('a:left/a:color/@indexed',ns) || na).value();

  this.diagonal = (borderStyles[id].get('a:diagonal',ns).attr('style') || na).value();
  this.diagonalColor = (borderStyles[id].get('a:diagonal/a:color/@indexed',ns) || na).value();
};

module.exports = Borders;
