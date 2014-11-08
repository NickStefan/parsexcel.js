var ns = require('../cell/namespace');
var getHex = require('./gethex');

var na = { 
  value: function() { return ''; },
  text:  function() { return ''; }
};

var Fills = function(id, fillStyles){
  this.patternFill = (fillStyles[id].get('a:patternFill',ns).attr('patternType') || na).value();
  if (this.patternFill === 'none') {
    this.patternFill = '';
  }
  this.foregroundColor = getHex((fillStyles[id].get('a:patternFill/a:fgColor/@indexed',ns) || na).value());
  this.backgroundColor = getHex((fillStyles[id].get('a:patternFill/a:bgColor/@indexed',ns) || na).value());
};

module.exports = Fills;
