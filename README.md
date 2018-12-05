# Excel template engine

The Documentation available on languages:
1. [English](https://github.com/optimistex/xlsx-template-ex#readme)
2. [Русский](README.ru.md)

Supported files: **xlsx**

The component implements making some template-based reports.
It has the flexible syntax similar to template expressions in the Angular framework. 

## The syntax

Supported 2 types of expressions:
* Output a property value: `{{value|pipe:arg1:argN}}`
* Output an array data: `[[value|pipe:arg1:argN]]`

Where:
* `value` - name of property that contain some value or an array
* `pipe` - some function for additional processing some value
* `arg1`, `argN` - arguments/parameters passing to a pipe function

## Implemented expression variants

* `{{propertyName}}` - output a value as is
* `{{propertyName|date}}` - the value formatted as date (DD.MM.YYYY)
* `{{propertyName|time}}` - the value formatted as time (hh:mm:ss) 
* `{{propertyName|datetime}}` - the value formatted as date and time (DD.MM.YYYY hh:mm:ss) 
* `{{fileName|image}}` - find a picture file by file name. 
    If the picture found, then it embed into a table cell 
* `{{propertyArrayName|find:propertyName:propertyValue}}` - find a value in the array `propertyArrayName` 
    that has the property `propertyName` that equal `propertyValue`
* `{{propertyObjectName|get:propertyName}}` - return a value of the property `propertyArrayName` from the object `propertyObjectName`     

* `[[array|repeat-rows:3]]` - process the array of values and output the content 
    into the section from 3 rows started from current.
    The rows will be duplicated according to the size of the array.
* `[[array|filter:propertyName:checkValue]]` - filter the array. 
    If provided only `propertyName`, then We will get an array of objects that contain the property.
    If provided `propertyName` and `checkValue`, then We will get an array of objects that contain 
    the property `propertyName` with value `checkValue`.
* `[[array|tile:blockRows:blockColumns:tileColumns]]` - process the array of values and output the data by blocks.
    The source block defines by `blockRows` and `blockColumns`. 
    The block will be output in a grid with `tileColumns` number columns. 
    
## Examples

We will output this data:
```javascript
let data = {
    reportBuildDate: 1526443275041,

    results: [
        { text: 'some text 1', answerText: 'a text of an answer 1'},
        { text: 'some text 2', answerText: 'a text of an answer 2'},
        { text: 'some text 3', answerText: 'a text of an answer 3'},
        { answerText: 'a text of an answer 3'},
    ],
};
```
    
Let's make a template:

**!!!** In the example / used instead of | because of a trouble with the markdown syntax 

| A | B |
|---|---|
|{{reportBuildDate/date}}| {{results/find:text:some text 2/get:answerText}} |
|[[results/filter:text/repeat-rows:1]] {{text}}| {{answerText}} |
| | |
| [[results/filter:text/tile:1:1:2]]{{answerText}} | |

Received result:

| A     | B     |
|-------|-------|
| 16.05.2018 | a text of an answer 2 |
| some text 1 | a text of an answer 1 |
| some text 2 | a text of an answer 2 |
| some text 3 | a text of an answer 3 |
| | |
| a text of an answer 1 | a text of an answer 2 |
| a text of an answer 3 | |

## Get started

```javascript
const fs = require("fs");
const XlsxTemplate = require('xlsx-template-ex');

const data = {
    reportBuildDate: 1526443275041,

    results: [
        { text: 'some text 1', answerText: 'a text of an answer 1'},
        { text: 'some text 2', answerText: 'a text of an answer 2'},
        { text: 'some text 3', answerText: 'a text of an answer 3'},
        { answerText: 'a text of an answer 3'},
    ],
};

XlsxTemplate.xlsxBuildByTemplate(data, 'template-file.xlsx')
    .then((buffer) => fs.writeFileSync('./out.xlsx', buffer))
    .catch((error) => console.log('xlsxHelper error:', error));
```

## Troubleshooting

Please follow this guidelines when reporting bugs and feature requests:

1. Use [GitHub Issues](https://github.com/optimistex/xlsx-template-ex/issues) board to report bugs and feature requests (not our email address)
2. Please **always** write steps to reproduce the error. That way we can focus on fixing the bug, not scratching our heads trying to reproduce it.

Thanks for understanding!

## Contribute

- `npm start` - Run the demo from a production build.
- `npm run start-dev` - Run the demo in a developing mode.
- `npm run build` - Build the demo for production.

# License

The MIT License (see the [LICENSE](https://github.com/optimistex/xlsx-template-ex/blob/master/LICENSE) file for the full text)
