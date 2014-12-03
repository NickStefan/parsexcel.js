[![NPM](https://nodei.co/npm/parsexcel.js.png?downloads=true&downloadRank=true&stars=true)](https://nodei.co/npm/parsexcel.js/)

parsexcel.js
========

Entire excel workbook parsed in native node.js.
 * Workbook meta data such as worksheet names, positions and dimensions.
 * Cell format such as general, date, boolean, currency, percent etc.
 * Cell styling information such as font size, alignment, style, and color.
 * Cell fill patterns and fill colors.
 * Cell borders and border styles.
 * Cell formulas, and calculated values.

Install
=======
    npm install parsexcel.js

NOTE: 0.0.3 API still supported for now, despite the documented API for 0.1.0 being different. The 0.0.3 API still exists in the form of parsexcel.parseRaw(), and should be favored over invoking parsexcel() itself (as that will likely be deprecated in the future).

parsexcel#htmlReady
==================
```
var parsexcel = require('parsexcel.js');

parsexcel.htmlReady('Spreadsheet.xlsx', function(err, data) {
  if(err) throw err;
    // ... do stuff with data ...
    // data is an object -- see Output below for documentation
});
```

parsexcel#htmlReady Output
==================

The output is one giant object. Full example result objects example can be found in test/answers.js. Here is the breakdown from outtermost to inner most properties:

## WORKSHEETS:
```
{
  1: {
    data: [ /* ... sheet data in rows ... */ ],
    sheetName: "The First Sheet",
    sheetPosition: "1",
  },
  2: {
    data: [ /* ... sheet data in rows ... */ ],
    sheetName: "The Second Sheet",
    sheetPosition: "2",
  }
}
```

## WORKSHEET ARRAY OF ROWS
The data property of each sheet is an array of row arrays. Each row array contains an index for each cell needed to represent the spread sheet. Each cell is represented as either a one space string `" "` or an object representing the cell:
```
data: [
  [ { /* cell object for A1 */ }, " ", " ", { /* cell object for A4 */ } ], /* row 1 */
  [ " ", " ", { /* cell object for B3 */ }, " ", { /* cell object for B5 */ } ] /* row 2 */
]
```

## CELL OBJECT
The cell object has value, formula, type, and style properties:
```
{
  value: "5",
  formula: "",
  type: "general",
  style: { /* ... style object ... */}
}
```
### Value and Cell Type
Value is always represented as a string, but when read with the type property can represent a number of cell types:

##### Date
```
value: "Sun Oct 20 2002 00:00:00 GMT-0700 (PDT)",
type: "date"
```

##### String
```
value: "bob",
type: "string"
```

##### General
```
value: "5",
type: "general"
```

##### Number
```
value: "5",
type: "number"
```

##### Currency
```
value: "100.05",
type: "currency"
```

##### Boolean
```
value: "0",
type: "boolean"
```

##### Percent
```
value: "0.97499999999999998",
type: "percent"
```

### Value and Formula

##### Basic Formulas
```
value: "10",
"formula": "SUM(B1:C1)"
```

##### Multi Sheet Formulas
```
value: "bobfromsheet2",
"formula": "Sheet2!A1"
```

## CELL STYLE OBJECT

Default values are empty strings: "". 
Otherwise, raw excel styles are
converted into something that can be easily added to a raw HTML node using camel cased javascript styles:

```
{
  // borders
  borderBottom: "", // example: thick solid
  borderBottomColor: "", // example: #000000
  borderTop: "",
  borderTopColor: "",
  borderRight: "",
  borderRightColor: "",
  borderLeft: "",
  borderLeftColor: "",

  // fill
  backgroundColor: "", // example: #000000

  // fonts
  fontFamily: ""; // example: "sans-serif"
  fontSize: "", // example: "12px"
  color: "", // example: #000000
  fontStyle: "", // example: "italic",
  textDecoration: "", // example: "underline",
  fontWeight: "", // example: "bold",

  // alignment
  verticalAlign: "", // example: "top-text",
  textAlign: ""; // example: "middle",
  
  // see the description of this object in parseRaw output
  format: { /* ... format object */ ... }
}
```




parsexcel#parseRaw
====
```
var parsexcel = require('parsexcel.js');

parsexcel.parseRaw('Spreadsheet.xlsx', function(err, data) {
  if(err) throw err;
    // ... do stuff with data ...
    // data is an object -- see Output below for documentation
});
```

parsexcel#parseRaw Output
======
This output is the exact same as the htmlReady output, except that the cell style object is closer to exactly what excel has stored:

## CELL STYLE OBJECT

```
"style": {
  "border": {
    "top": "", /* thick, thin etc, or "" */
    "topColor": "", /* hex value or "" */
    "bottom": "",
    "bottomColor": "",
    "right": "",
    "rightColor": "",
    "left": "",
    "leftColor": "",
    "diagonal": "",
    "diagonalColor": ""
  },
  "fill": { /* ... fill object ... */ },
  "font": {
    "size": "10",
    "bold": false,
    "italic": false,
    "underlined": false,
    "name": "Verdana",
    "color": "" /* hex value if applicable */
  },
  "alignment": {
    "vertical": "", /* top, bottom, center, or ""  */
    "horizontal": "" /* left, center, right, justified, or "" */
  },
  "format": { /* ... format object */ ... }
}
```

## FILL OBJECT

### Cell with no fill
```
"fill": {
  "patternFill": "",
  "foregroundColor": "",
  "backgroundColor": ""
}
```

### Cell with basic fill
```
"fill": {
  "patternFill": "solid",
  "foregroundColor": "#FF9900", /* actual color hex value that matters */
  "backgroundColor": "#000000" /* value can be ignored */
}
```

### Cell with patterned shading of two different colors
```
"fill": {
  "patternFill": "mediumGray", /* dotted shading pattern to foreground color */
  "foregroundColor": "#FFFF00",/* color of shading dots */
  "backgroundColor": "#00FFFF" /* color behind the dot pattern */
}
```

## FORMAT OBJECT
There are 180+ excel cell formats! There is currently an object look up for about 70 of them. The missing ones are either esoteric, eastern language, or very legacy.

Here's a quick tour of the 5 types. Each type has many different supported string formats.

##### General
```
"format": {
  "string": "General",
  "type": "general"
}
```

##### Number
```
"format": {
  "string": "General",
  "type": "general"
}
```

##### Currency
```
"format": {
  "string": "($#,##0_);[Red]($#,##0)",
  "type": "currency"
}
```

##### Percent
```
"format": {
  "string": "0.00%",
  "type": "percent"
}
```

##### Date
```
"format": {
  "string": "mm-dd-yy",
  "type": "date"
}
```

There are 65 other ones with different string formats that are not listed here, but can be found in /cellstyle/numformats.js


Backwards Compatability
=======================
In order to maintain backwards compatability, the old API (which can currently be used with parsexcel.parseRaw() ) is still available as:
```
var parsexcel = require('parsexcel.js');

parsexcel('Spreadsheet.xlsx', function(err, data) {
  if(err) throw err;
    // ... do stuff with data ...
    // data is an object -- see parseRaw Output for documentation
});
```

Test
=====
    npm test

Inspect Excel's XML
=======
    npm run inspect-xlsx      

This will run a bash script to unzip the test folder xlsx files into xml files. Editing a parsing library sometimes requires inspecting the XML files that you're trying to parse.

MIT License.

**parsexcel.js was originally a fork of [excel.js](https://github.com/trevordixon/excel.js). It was re-released as a different project because the code base is almost 80% divergent from excel.js. I am however greatful for the dimensions calculations and the skeleton provided by the original repo. Thank you to [trevordixon](https://github.com/trevordixon/), and other contributors, for excel.js.**