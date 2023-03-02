# [Excelsheet Handler]

The Excelsheet Handler offers battle-tested open-source solutions for converting excelsheet 
buffer data to json data and generating new spreadsheets from JSON array data, that will work
with legacy and modern software alike.

## Getting Started

### Installation

**NodeJS**

With [npm](https://www.npmjs.org/package/alan-excel-handler):

```bash
$ npm install alan-excel-handler
```

By default, the module supports `require`:

```js
var ExcelHandler = require("alan-excel-handler");
```

To Convert JSON Array data, and Export Excelsheet : (e.g.)

```js
var ExcelHandler = require("alan-excel-handler");

let path = `uploads/sample.xlsx`;
const info = {path: path};
const excelHandler = new ExcelHandler(info);
try {
    const users = [
        { name: "John", age: 25, email: "john@example.com" },
        { name: "Mary", age: 30, email: "mary@example.com" },
        { name: "Jane", age: 35, email: "jane@example.com" }
    ];

    const config = {
        arrData: users,
        sheetName: "sampleSheet"
    };
    await jsonConverter.exportSheets(config);
}catch (err) {
    console.log(`err: `, err);
    throw err;
}
```
