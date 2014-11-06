var Promise = require('bluebird'),
	_ = require('underscore');

function extractFiles(path,cb) {
	var AdmZip = require('adm-zip');
  
  // desired files
	var files = {
		'xl/worksheets/sheet1.xml': {},
		'xl/sharedStrings.xml': {},
		'xl/styles.xml': {}
	};

  // reading archives
  var zip = new AdmZip(path);
  var zipEntries = zip.getEntries();

  zipEntries.forEach(function(entry,i,c) {
      if (files[entry.entryName]) {
        files[entry.entryName].contents = zip.readAsText(entry);
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
	try {
		var libxmljs = require('libxmljs'),
			sheet = libxmljs.parseXml(files['xl/worksheets/sheet1.xml'].contents),
			ns = {a: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'},
			data = [],
			strings,
			styles;
      
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


	} catch(parseError){
	   return [];
  }

	var colToInt = function(col) {
		var letters = ["", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"];
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

	var cellNodes = sheet.find('/a:worksheet/a:sheetData/a:row/a:c', ns);
	var cells = _(cellNodes).map(function (node) {
		return new Cell(node);
	});

  var cellStyles = styles.find('/a:styleSheet/a:cellXfs/a:xf', ns);
	var styleMap = _(cellStyles).map(function (node) {
		return new CellStyle(node);
	});

	var d = sheet.get('//a:dimension/@ref', ns);
	if (d) {
		d = _.map(d.value().split(':'), function(v) { return new CellCoords(v); });
	} else {
        d = calculateDimensions(cells)
	}

	var cols = d[1].column - d[0].column + 1,
		rows = d[1].row - d[0].row + 1;

	_(rows).times(function() {
		var _row = [];
		_(cols).times(function() { _row.push(''); });
		data.push(_row);
	});
  
	_.each(cells, function(cell) {
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

		data[cell.row - d[0].row][cell.column - d[0].column] = value;
	});
	return data;
}

module.exports = function parseXcel(path, cb) {
  extractFilesSync(path)
		.then(function(files) {
			cb(null, extractData(files));
		});
};
