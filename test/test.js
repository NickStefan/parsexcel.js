var parseXcel = require('../parseXcel');
var assert = require('assert');
var answers = require('./answers');

var sheetsDir = __dirname + '/spreadsheets';

describe('excel.js', function() {
  for (var filename in answers) {
    (function(filename, expected) {
    
      if (filename === 'excel_mac_2008-numbers.xlsx'){
        describe(filename + ' test', function() {
          it('should return the right numbers and formatting', function(done) {
            parseXcel(sheetsDir + '/' + filename, function(err, data) {
              console.log(JSON.stringify(data,null,2));
              assert.deepEqual(data, expected);
              done(err);
            });
          });
        });
        
      } else if (filename === 'excel_mac_2008-formula.xlsx'){
        describe(filename + ' basic test', function() {
          it('should return the right formulas and formatting', function(done) {
            parseXcel(sheetsDir + '/' + filename, function(err, data) {
              console.log(JSON.stringify(data,null,2));
              assert.deepEqual(data, expected);
              done(err);
            });
          });
        });

      } else if (filename === 'excel_mac_2011-formatting.xlsx'){
        describe(filename + ' test', function() {
          it('should return the right formatting and multi-sheet workbook', function(done) {
            parseXcel(sheetsDir + '/' + filename, function(err, data) {
              console.log(JSON.stringify(data,null,2));
              assert.deepEqual(data, expected);
              done(err);
            });
          });
        });
      }

    })(filename, answers[filename]);
  }
});
