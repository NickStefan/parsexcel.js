var Promise = require('bluebird'),
	_ = require('underscore');

function extractFiles(path,cb) {
	var AdmZip = require('adm-zip');
  
  // desired files
	var files = {
		'xl/worksheets/sheet1.xml': {},
		'xl/sharedStrings.xml': {}
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
			strings;
      
      // if excel file has no strings, sharedStrings won't exist
      if (files['xl/sharedStrings.xml'].contents){
			  strings = libxmljs.parseXml(files['xl/sharedStrings.xml'].contents);
			} else {
        strings = false;
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
			coords = new CellCoords(r);

		this.column = coords.column;
		this.row = coords.row;
		this.value = value;
		this.formula = formula;
		this.type = type;
	};

	var cellNodes = sheet.find('/a:worksheet/a:sheetData/a:row/a:c', ns);
	var cells = _(cellNodes).map(function (node) {
		return new Cell(node);
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
  
  if (options.getAll) {
  	_.each(cells, function(cell) {
			var value = {
				v: cell.value,
				f: cell.formula
			};

			if (cell.type == 's') {
				values = strings.find('//a:si[' + (parseInt(value.v) + 1) + ']//a:t[not(ancestor::a:rPh)]', ns)
				value.v = "";
				for (var i = 0; i < values.length; i++) {
					value.v += values[i].text();
				}
			}
			data[cell.row - d[0].row][cell.column - d[0].column] = value;
		});

  } else {
		_.each(cells, function(cell) {
			var value = cell.value

			if (cell.type == 's') {
				values = strings.find('//a:si[' + (parseInt(value) + 1) + ']//a:t[not(ancestor::a:rPh)]', ns)
				value = "";
				for (var i = 0; i < values.length; i++) {
					value += values[i].text();
				}
			}
			data[cell.row - d[0].row][cell.column - d[0].column] = value;
		});
	}
	return data;
}

module.exports = function parseXlsx(path, cb) {
  extractFilesSync(path)
		.then(function(files) {
			cb(null, extractData(files,{getAll:false}));
		});
};

module.exports.all = function all(path, cb) {
  extractFilesSync(path)
		.then(function(files) {
			cb(null, extractData(files,{getAll:true}));
		});
};
