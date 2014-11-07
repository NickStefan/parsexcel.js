var _ = require('underscore');
var Promise = require('bluebird');
var AdmZip = require('adm-zip');
var libxmljs = require('libxmljs');

var Cell = require('./cell');
var CellCoords = require('./cellcoords');
var CellStyle = require('./cellstyle');
var calcSheetDimensions = require('./calcsheetdimensions');
var ns = {a: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'};


function extractFiles(path,cb) {
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

  zipEntries.forEach(function(entry) {
    // filename in our wanted list
    if (files[entry.entryName]) {
      files[entry.entryName].contents = zip.readAsText(entry);
    // filename is one of the worksheets
    } else if (/xl\/worksheets\/sheet/.test(entry.entryName)) {
      files.sheets[entry.entryName] = {};
      files.sheets[entry.entryName].contents = zip.readAsText(entry);
    }
  });
  cb(files);
}

function extractData(files) {
  var sheets = {};
	var output = {};
	var strings;
	var styles;
  var workbook;

  for (var sheet in files.sheets){
    // data to parse
    sheets[sheet] = {};
    sheets[sheet].data = libxmljs.parseXml(files.sheets[sheet].contents);
    // output object to eventually return
    output[sheet] = {};
    output[sheet].data = [];
  }

  workbook = libxmljs.parseXml(files['xl/workbook.xml'].contents);
  
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
  
  var wbSheetInfo = {};

  workbook.find('/a:workbook/a:sheets/a:sheet', ns)
    .forEach(function(node){
      wbSheetInfo[node.attr('id').value().replace(/rId/,'')] = {
        name: node.attr('name').value(),
        position: node.attr('id').value().replace(/rId/,'')
      }
    });
    

  _(sheets).each(function(sheet,sheetName){
    // get the sheet cells
    sheet.cells = sheet.data.find('/a:worksheet/a:sheetData/a:row/a:c', ns)
      .map(function(node){
        return new Cell(node);
      });
    
    // get the sheet dimensions
    calcSheetDimensions(sheet);

    var cols = sheet.dimensions[1] ? sheet.dimensions[1].column - sheet.dimensions[0].column + 1 : 1;
    var rows = sheet.dimensions[1] ? sheet.dimensions[1].row - sheet.dimensions[0].row + 1 : 1;
    
    // add the number of array elements to match the sheet dimensions
    _(rows).times(function() {
      var _row = [];
      _(cols).times(function() { _row.push(''); });
      output[sheetName].data.push(_row);
    });
  
  });
  
  var rawStyles = {};
  rawStyles.fontStyles = styles.find('/a:styleSheet/a:fonts/a:font', ns);
  rawStyles.fillStyles = styles.find('/a:styleSheet/a:fills/a:fill', ns);
  rawStyles.borderStyles = styles.find('/a:styleSheet/a:borders/a:border', ns);

  var styleMap = styles.find('/a:styleSheet/a:cellXfs/a:xf', ns)
	  .map(function (node) {
  		return new CellStyle(node,rawStyles);
  	});
  
	_(sheets).each(function(sheet,sheetFileName) {
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

  		output[sheetFileName].data[cell.row - sheet.dimensions[0].row][cell.column - sheet.dimensions[0].column] = value;
    });
    
    // update the individual sheet meta data
    var sheetId = sheetFileName.replace(/xl\/worksheets\/sheet|.xml/g,'');
    output[sheetFileName].sheetName = wbSheetInfo[sheetId].name;
    output[sheetFileName].sheetPosition = wbSheetInfo[sheetId].position;
    output[sheetId] = output[sheetFileName];
    delete output[sheetFileName];
	});
	return output;
}

function extractFilesSync(path){
  return new Promise(function(resolve,reject){
    extractFiles(path,function(data){
      resolve(data);
    });
  });
}

module.exports = function parseXcel(path, cb) {
  extractFilesSync(path)
		.then(function(files) {
			cb(null, extractData(files));
		});
};
