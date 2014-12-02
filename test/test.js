var parsexcel = require('../parsexcel');
var assert = require('assert');
assert.deepEqualWithDifflet = require('deep-equal-with-difflet');
var answers = require('./answers');

var sheetsDir = __dirname + '/spreadsheets';

describe('should parse raw excel data', function() {
  for (var filename in answers.parseRaw) {
    (function(filename, expected) {
    
      if (filename === 'excel_mac_2008-types.xlsx'){
        describe(filename + ' test \n', function() {
          it('should return correct numbers, currencies, strings, booleans, and dates', function(done) {
            parsexcel.parseRaw(sheetsDir + '/' + filename, function(err, data) {
              //console.log(JSON.stringify(data,null,2));
              assert.deepEqualWithDifflet(data, expected);
              done(err);
            });
          });
        });
        
      } else if (filename === 'excel_mac_2008-formulas.xlsx'){
        describe(filename + ' test \n', function() {
          it('should return the correct formulas, even accross spreadsheets', function(done) {
            parsexcel.parseRaw(sheetsDir + '/' + filename, function(err, data) {
              //console.log(JSON.stringify(data,null,2));
              assert.deepEqualWithDifflet(data, expected);
              done(err);
            });
          });
        });

      } else if (filename === 'excel_mac_2011-formats.xlsx'){
        describe(filename + ' test \n', function() {
          it('should return the correct cell formatting and styles', function(done) {
            parsexcel.parseRaw(sheetsDir + '/' + filename, function(err, data) {
              //console.log(JSON.stringify(data,null,2));
              assert.deepEqualWithDifflet(data, expected);
              done(err);
            });
          });
        });
      }

    })(filename, answers.parseRaw[filename]);
  }
});

describe('should parse htmlReady excel data and style data', function() {
  for (var filename in answers.htmlReady) {
    (function(filename, expected) {
      if (filename === 'excel_mac_2011-formats.xlsx'){
        describe(filename + ' test \n', function() {
          it('should return correct numbers, currencies, strings, booleans, and dates', function(done) {
            parsexcel.htmlReady(sheetsDir + '/' + filename, function(err, data) {
              //console.log(JSON.stringify(data,null,2));
              assert.deepEqualWithDifflet(data, expected);
              done(err);
            });
          });
        });
      }

    })(filename, answers.htmlReady[filename]);
  }
});

