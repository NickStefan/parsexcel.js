var ns = require('../cell/namespace');
var getHex = require('./gethex');

var na = { 
  value: function() { return ''; },
  text:  function() { return ''; }
};

var Fonts = function(id, fontStyles){
  this.size = (fontStyles[id].get('a:sz', ns).attr('val') || na).value();
  this.bold = (fontStyles[id].get('a:b', ns) ? true : false);
  this.italic = (fontStyles[id].get('a:i', ns) ? true : false);
  this.underlined = (fontStyles[id].get('a:u', ns) ? true : false);
  this.name = (fontStyles[id].get('a:name', ns).attr('val') || na).value();
  this.color = getHex((fontStyles[id].get('a:color/@indexed', ns) || na).value());
};

module.exports = Fonts;
