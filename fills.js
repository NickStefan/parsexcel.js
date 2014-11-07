var na = { 
  value: function() { return ''; },
  text:  function() { return ''; }
};

var ns = {a: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'};

var Fills = function(id, fillStyles){
  this.patternFill = (fillStyles[id].get('a:patternFill',ns).attr('patternType') || na).value();
  if (this.patternFill === 'none') {
    this.patternFill = '';
  }
  this.foregroundColor = (fillStyles[id].get('a:patternFill/a:fgColor/@indexed',ns) || na).value();
  this.backgroundColor = (fillStyles[id].get('a:patternFill/a:bgColor/@indexed',ns) || na).value();
};

module.exports = Fills;
