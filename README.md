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

To Convert JSON Array data, and Export Excelsheet (For Single Sheet): (e.g.)

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
        sheetName: "users"
    };
    await excelHandler.exportSheets(config);
}catch (err) {
    console.log(`err: `, err);
    throw err;
}
```

To Convert JSON Array data, and Export Excelsheet (For Multiple Sheets): (e.g.)

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
    
    const students = [
        { rollNo: 1, name: "Alice", grade: 85 },
        { rollNo: 2, name: "Bob", grade: 92 },
        { rollNo: 3, name: "Charlie", grade: 70 },
        { rollNo: 4, name: "David", grade: 88 }
    ];

    const config1 = {
        arrData: users,
        sheetName: "users",
        isMultipleSheets: true
    };
    await excelHandler.exportSheets(config1);

    const config2 = {
        arrData: students,
        sheetName: "students",
        isLastSheet: true,
        isMultipleSheets: true
    }
    await excelHandler.exportSheets(config2);

}catch (err) {
    console.log(`err: `, err);
    throw err;
}
```