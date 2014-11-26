
// I am definitely not 100% confident on these types!
// I am sure on the string formats corresponding to the codes
var formats = {
  0 :['general','General'],
  1 :['number','0'],
  2 :['number','0.00'],
  3 :['number','#,##0'],
  4 :['number','#,##0.00'],
  5 :['currency','($#,##0_);($#,##0)'],
  6 :['currency','($#,##0_);[Red]($#,##0)'],
  7 :['currency','($#,##0.00_);($#,##0.00)'],
  8 :['currency','($#,##0.00_);[Red]($#,##0.00)'],
  9 :['percent','0%'],
  10:['percent','0.00%'],
  11:['number','0.00E+00'],
  12:['number','# ?/?'],
  13:['number','# ??/??'],
  14:['date','mm-dd-yy'],
  15:['date','d-mmm-yy'],
  16:['date','d-mmm'],
  17:['date','mmm-yy'],
  18:['date','h:mm AM/PM'],
  19:['date','h:mm:ss AM/PM'],
  20:['date','h:mm'],
  21:['date','h:mm:ss'],
  22:['date','m/d/yy h:mm'],

  27:['date','[$-404]e/m/d'],
  30:['date','m/d/yy'],

  36:['date','[$-404]e/m/d'],
  37:['currency','#,##0 ;(#,##0)'],
  38:['currency','#,##0 ;[Red](#,##0)'],
  39:['currency','#,##0.00;(#,##0.00)'],
  40:['currency','#,##0.00;[Red](#,##0.00)'],
  41:['currency','_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)'],
  42:['currency','_($* #,##0_);_($* (#,##0);_($* "-"_);_(@_)'],
  43:['currency','_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'],
  44:['currency','_("$"* #,##0.00_);_("$"* \(#,##0.00\);_("$"* "-"??_);_(@_)'],
  45:['date','mm:ss'],
  46:['date','[h]:mm:ss'],
  47:['date','mmss.0'],
  48:['number','##0.0E+0'],
  49:['number','@'],

  
  50:['date','[$-404]e/m/d'],
  57:['date','[$-404]e/m/d'],

  59:['number','t0'],
  60:['number','t0.00'],
  61:['number','t#,##0'],
  62:['number','t#,##0.00'],
  67:['number','t0%'],
  68:['number','t0.00%'],
  69:['number','t# ?/?'],
  70:['number','t# ??/??'],

  165:['number','????'],
  166:['number','????'],
  167:['number','????'],
  168:['number','????'],
  169:['number','????'],
  170:['number','????'],
  171:['number','????'],
  172:['number','????'],
  173:['number','????'],
  174:['number','????'],
  175:['number','????'],
  176:['number','????'],
  177:['number','????'],
  178:['number','????'],
  179:['number','????'],
  180:['number','????'],
  181:['number','????'],
  182:['number','????'],
  183:['number','????'],
  184:['number','????'],
  185:['number','????'],
  186:['number','????'],
};

var NumFormats = function(id){
  this.string = formats[id] !== undefined ? formats[id][1] : 'General';
  this.type = formats[id] !== undefined ? formats[id][0] : 'general';
};

module.exports = NumFormats;

