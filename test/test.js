var parseXcel = require('../parseXcel');
var assert = require('assert');
var answers = require('./answers');

var sheetsDir = __dirname + '/spreadsheets';

describe('excel.js', function() {
  for (var filename in answers) {
    (function(filename, expected) {
    
      if (filename === 'excel_mac_2008-types.xlsx'){
        describe(filename + ' test', function() {
          it('should return correct numbers, currencies, strings, booleans, and dates', function(done) {
            parseXcel(sheetsDir + '/' + filename, function(err, data) {
              console.log(JSON.stringify(data,null,2));
              console.log(JSON.stringify(expected,null,2));
              assert.deepEqual(data, expected);
              done(err);
            });
          });
        });
        
      } else if (filename === 'excel_mac_2008-formulas.xlsx'){
        describe(filename + ' basic test', function() {
          it('should return the correct formulas', function(done) {
            parseXcel(sheetsDir + '/' + filename, function(err, data) {
              //console.log(JSON.stringify(data,null,2));
              assert.deepEqual(data, expected[filename]);
              done(err);
            });
          });
        });

      } else if (filename === 'excel_mac_2011-formats.xlsx'){
        describe(filename + ' test', function() {
          it('should return the correct cell formatting and styles', function(done) {
            parseXcel(sheetsDir + '/' + filename, function(err, data) {
              //console.log(JSON.stringify(data,null,2));
              assert.deepEqual(data, expected[filename]);
              done(err);
            });
          });
        });
      }

    })(filename, answers[filename]);
  }
});
