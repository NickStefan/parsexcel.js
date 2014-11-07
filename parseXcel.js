var Promise = require('bluebird'),
	_ = require('underscore');

function extractFiles(path,cb) {
	var AdmZip = require('adm-zip');
  
  // desired files in addition to sheets
	var files = {
		'xl/sharedStrings.xml': {},
		'xl/styles.xml': {},
    'xl/workbook.xml':{},
    sheets:{}
	};
  var sheets = {};

  // reading archives
  var zip = new AdmZip(path);
  var zipEntries = zip.getEntries();

  zipEntries.forEach(function(entry,i,c) {
      if (files[entry.entryName]) {
        files[entry.entryName].contents = zip.readAsText(entry);
      } else if (/xl\/worksheets\/sheet/.test(entry.entryName)) {
        files.sheets[entry.entryName] = {};
        files.sheets[entry.entryName].contents = zip.readAsText(entry);
      }
  });
  cb(files);
}

function extractFilesSync(path){
	return new Promise(function(resolve,reject){
		extractFiles(path,function(data){
      resolve(data);
		});
	});
}

function calculateDimensions (cells) {
    var comparator = function (a, b) { return a-b; };
    var allRows = _(cells).map(function (cell) { return cell.row; }).sort(comparator),
        allCols = _(cells).map(function (cell) { return cell.column; }).sort(comparator),
        minRow = allRows[0],
        maxRow = _.last(allRows),
        minCol = allCols[0],
        maxCol = _.last(allCols);

    return [
        {row: minRow, column: minCol},
        {row: maxRow, column: maxCol}
    ];
}

function extractData(files,options) {
	var libxmljs = require('libxmljs');
  var sheets = {};
	var ns = {a: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'};
	var output = {};
	var strings;
	var styles;

    for (var sheet in files.sheets){
      // data to parse
      sheets[sheet] = {};
      sheets[sheet].data = libxmljs.parseXml(files.sheets[sheet].contents);
      // output object to eventually return
      output[sheet] = {};
      output[sheet].data = [];
    }
    
    // if excel file has no strings, sharedStrings won't exist
    if (files['xl/sharedStrings.xml'].contents){
		  strings = libxmljs.parseXml(files['xl/sharedStrings.xml'].contents);
		} else {
      strings = false;
    }
    
    // if excel file has no styles, styles won't exist
    if (files['xl/styles.xml'].contents){
		  styles = libxmljs.parseXml(files['xl/styles.xml'].contents);
		} else {
      styles = false;
    }

	var colToInt = function(col) {
		var letters = [
      "", "A", "B", "C", "D", "E", "F", "G", 
      "H", "I", "J", "K", "L", "M", "N", "O", "P",
       "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"
    ];
		col = col.trim().split('');
		
		var n = 0;

		for (var i = 0; i < col.length; i++) {
			n *= 26;
			n += letters.indexOf(col[i]);
		}

		return n;
	};

	var CellCoords = function(cell) {
		cell = cell.split(/([0-9]+)/);
		this.row = parseInt(cell[1]);
		this.column = colToInt(cell[0]);
	};

	var na = { 
		value: function() { return ''; },
    text:  function() { return ''; }
  };

	var Cell = function(cellNode) {
		var r = cellNode.attr('r').value(),
			type = (cellNode.attr('t') || na).value(),
			value = (cellNode.get('a:v', ns) || na ).text(),
			formula = (cellNode.get('a:f', ns) || na).text(),
			style = (cellNode.attr('s') || na).value(),
			coords = new CellCoords(r);

		this.column = coords.column;
		this.row = coords.row;
		this.value = value;
		this.formula = formula;
		this.type = type;
		this.style = style;
	};

  
	var fontStyles = styles.find('/a:styleSheet/a:fonts/a:font', ns);
  var fillStyles = styles.find('/a:styleSheet/a:fills/a:fill', ns);
  var borderStyles = styles.find('/a:styleSheet/a:borders/a:border', ns);

  var Borders = function(id){
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
  var Fills = function(id){
    this.patternFill = (fillStyles[id].get('a:patternFill',ns).attr('patternType') || na).value();
    if (this.patternFill === 'none') {
    	this.patternFill = '';
    }
    this.foregroundColor = (fillStyles[id].get('a:patternFill/a:fgColor/@indexed',ns) || na).value();
    this.backgroundColor = (fillStyles[id].get('a:patternFill/a:bgColor/@indexed',ns) || na).value();
  };
  var Fonts = function(id){
    this.size = (fontStyles[id].get('a:sz', ns).attr('val') || na).value();
    this.bold = (fontStyles[id].get('a:b', ns) ? true : false);
    this.italic = (fontStyles[id].get('a:i', ns) ? true : false);
    this.underlined = (fontStyles[id].get('a:u', ns) ? true : false);
    this.name = (fontStyles[id].get('a:name', ns).attr('val') || na).value();
    this.color = (fontStyles[id].get('a:color/@indexed', ns) || na).value();
  };
  var Alignments = function(cellNode){
  	this.vertical = (cellNode.get('a:alignment/@vertical', ns) || na).value();
    this.horizontal = (cellNode.get('a:alignment/@horizontal', ns) || na).value();
  };

	var CellStyle = function(cellNode) {
    var fontId = cellNode.attr('fontId').value(),
      fillId = cellNode.attr('fillId').value(),
      borderId = cellNode.attr('borderId').value();

		this.border = new Borders(borderId);
		this.fill = new Fills(fillId);
    this.font = new Fonts(fontId);
    this.alignment = new Alignments(cellNode);
	};
  
  _(sheets).each(function(v,k,c){
    // get the sheet cells
    sheets[k].cells = sheets[k].data.find('/a:worksheet/a:sheetData/a:row/a:c', ns)
      .map(function(node){
        return new Cell(node);
      });
    
    // get the sheet dimensions
    var d = sheets[k].data.get('//a:dimension/@ref', ns);
    if (d) {
      sheets[k].d = _.map(d.value().split(':'), function(v) { return new CellCoords(v); });
    } else {
      sheets[k].d = calculateDimensions(cells)
    }

    var cols = sheets[k].d[1] ? sheets[k].d[1].column - sheets[k].d[0].column + 1 : 1;
    var rows = sheets[k].d[1] ? sheets[k].d[1].row - sheets[k].d[0].row + 1 : 1;
    
    // add the number of array elements to match the sheet dimensions
    _(rows).times(function() {
      var _row = [];
      _(cols).times(function() { _row.push(''); });
      output[k].data.push(_row);
    });
  
  });

  var cellStyles = styles.find('/a:styleSheet/a:cellXfs/a:xf', ns);
	var styleMap = _(cellStyles).map(function (node) {
		return new CellStyle(node);
	});
  
	_(sheets).each(function(sheet,k,c) {
    _(sheet.cells).each(function(cell){
  		var value = {
  			value: cell.value,
  			formula: cell.formula,
  			type: cell.type,
  			style: styleMap[0]
  		};

  		if (cell.type === 's') {
  			values = strings.find('//a:si[' + (parseInt(value.value) + 1) + ']//a:t[not(ancestor::a:rPh)]', ns)
  			value.value = "";
  			for (var i = 0; i < values.length; i++) {
  				value.value += values[i].text();
  			}
  			value.type = 'string';
  		}
      
  		if (parseInt(cell.style) >= 0) {
  			value.style = styleMap[parseInt(cell.style)];
  		}

  		output[k].data[cell.row - sheet.d[0].row][cell.column - sheet.d[0].column] = value;
    });
	});
	return output;
}

module.exports = function parseXcel(path, cb) {
  extractFilesSync(path)
		.then(function(files) {
			cb(null, extractData(files));
		});
};
