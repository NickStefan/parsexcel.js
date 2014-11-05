var parseXlsx = require('../excelParser');
var assert = require('assert');

var sheetsDir = __dirname + '/spreadsheets';
var sheets = {
  'excel_mac_2008-numbers.xlsx': [ 
    [{v:'0', f:''},{v:'11',f:''},{v:'5',f:''}]
  ],

  'excel_mac_2008-formula.xlsx': [ 
      [{v:'10', f:'SUM(B1:C1)'},{v:'5',f:''},{v:'5',f:''}],
      [{v:'bob',f:''},'','']
    ],

  'excel_mac_2011-basic.xlsx': [
      [{v:'One',f:''},{v:'Two',f:''},{v:'10',f:'SUM(D1:E1)'},{v:'5',f:''},{v:'5',f:''}],
      [{v:'Three',f:''},{v:'Four',f:''},'','',''] 
    ],

  'excel_mac_2011-formatting.xlsx': [ 
      [{v:'Hey',f:''},{v:'now',f:''},{v:'so',f:''} ],
      [{v:'cool',f:''},'',''] 
    ]
};

describe('excel.js', function() {
  for (var filename in sheets) {
    (function(filename, expected) {

      describe(filename + ' basic test', function() {
        it('should return the right value', function(done) {
          parseXlsx.all(sheetsDir + '/' + filename, function(err, data) {
            assert.deepEqual(data, expected);
            done(err);
          });
        });
      });

    })(filename, sheets[filename]);
  }
});

var sheetsOldApi = {
  'excel_mac_2011-basic.xlsx': [ [ 'One', 'Two', '10','5','5'], [ 'Three', 'Four','','',''] ],
  'excel_mac_2011-formatting.xlsx': [ [ 'Hey', 'now', 'so' ], [ 'cool', '', '' ] ]
};

describe('excel.js', function() {
  for (var filename in sheetsOldApi) {
    (function(filename, expected) {

      describe(filename + ' basic test', function() {
        it('should return the right value', function(done) {
          parseXlsx(sheetsDir + '/' + filename, function(err, data) {
            assert.deepEqual(data, expected);
            done(err);
          });
        })
      });

    })(filename, sheetsOldApi[filename]);
  }
});
