ParseXcel.js
========

Entire excel workbook parsed in native node.js.
 * Workbook meta data such as worksheet names, positions and dimensions.
 * Cell styling information such as font size, alignment, style, and color.
 * Cell fill patterns and fill colors.
 * Cell borders and border styles.
 * Cell formulas, and calculated values.

Install
=======
    npm install ParseXcel.js

Use
====
    var parseXcel = require('excel');

    parseXcel('Spreadsheet.xlsx', function(err, data) {
      if(err) throw err;
        // data is an array of arrays of objects with cell properties
    });


Output
======

    
Test
=====
Run `npm test`

Inspect
=======
Inspect the xml files inside of the test/spreadsheets/*.xlsx files with:
`npm run inspect-xlsx`

This will run a bash script to unzip the test folder xlsx files into xml files. Check the contents folders inside of the test/spreadsheets/ folder after running.

MIT License.

**Project was originally a fork of [excel.js](https://github.com/trevordixon/excel.js). Thank you to [trevordixon](https://github.com/trevordixon/) for the original inspiration.**