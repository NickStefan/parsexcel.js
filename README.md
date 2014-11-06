ParseXcel.js
========

Entire excel workbook parsed in native node.js. Get styles, forumulas, values, etc.

Install
=======
    npm install ParseXcel.js

Use
====
    var parseXlsx = require('excel');

    parseXcel('Spreadsheet.xlsx', function(err, data) {
      if(err) throw err;
        // data is an array of arrays of objects with cell properties
    });
    
Test
=====
Run `npm test`

MIT License.

**Project was originally a fork of [excel.js](https://github.com/trevordixon/excel.js). Thank you to [trevordixon](https://github.com/trevordixon/) for the original inspiration.**