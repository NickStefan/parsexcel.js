
// convert parsexcel.js styles into js camelCase HTML styling formats
var htmlStyles = function(style){
  // borders
  this.borderBottom = addSolid(style.border.bottom);
  this.borderBottomColor = style.border.bottomColor;
  this.borderTop = addSolid(style.border.bottom);
  this.borderTopColor = style.border.bottomColor;
  this.borderRight = addSolid(style.border.bottom);
  this.borderRightColor = style.border.bottomColor;
  this.borderLeft = addSolid(style.border.bottom);
  this.borderLeftColor = style.border.bottomColor;

  // fill
  if (style.fill.foregroundColor && style.fill.patternFill === 'solid') {
    // this is not an error. its actualy backwards in excel
    this.backgroundColor = style.fill.foregroundColor;
  } else if (style.fill.foregroundColor) {
    // should mix colors somehow...
    // for now just set it to the background color
    this.backgroundColor = style.fill.backgroundColor;
  }

  // fonts
  this.fontFamily = style.font.name;
  this.fontSize = addPx(style.font.size);
  this.color = style.font.color;
  this.fontStyle = style.font.italic ? 'italic' : ''; // italic
  this.textDecoration = style.font.underlined ? 'underline' : ''; // underlined
  this.fontWeight = style.font.bold ? 'bold' : ''; // bold

  // alignment
  this.verticalAlign = align(style.alignment.vertical);
  this.textAlign = style.alignment.horizontal;

  // format key is deleted inside parsexcel before being outputted to the user
  // its built this way to deal with how the parser was originally written
  // this minimizes some of a possible refactor
  this.format = style.format;
}

function addSolid(val){
  if (val){
    return val + " solid";
  }
  return "";
}

function addPx(val){
  if (val){
    return val + "px";
  }
  return "";
}

function align(val){
  if (val === "center"){
    return "middle";
  }
  if (val === "top" || val === "top-text"){
    return "top-text";
  }
  if (val === "bottom" || val === "bottom-text"){
    return "bottom-text";
  }
  return "";
}

module.exports = htmlStyles;
