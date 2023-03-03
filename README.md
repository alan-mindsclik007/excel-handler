# Excelsheet Handler

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

_Excel Properties_

The `config` is an objects which have the following properties:

```typescript
    const config = {
        /* What are the rows you want to export ? */
        arrData: Array,
        /* What is you sheet name ? */
        sheetName: String, // Optional for single sheet
        /* Is this your last sheet ? */
        isLastSheet: Boolean, // Optional for single sheet
        /* Do you want generate Excelsheet with having Multiple Sheets ? */
        isMultipleSheets: Boolean  // if true, the more than one sheet, And Optional for single sheet
    };
```

To Convert JSON Array data, and Export Excelsheet (For Single Sheet): (e.g.)

```js
var ExcelHandler = require("alan-excel-handler");

let path = `uploads/sample.xlsx`;
const info = {path: path};
const excelHandler = new ExcelHandler(info);
try {
    const users = [
        { name: "Alankar Awadhiya", age: 27, email: "alanawadhiya007@gmail.com" },
        { name: "Atul", age: 25, email: "atulawadhiya@example.com" },
        { name: "Niharika", age: 35, email: "niharikaawadhiya@example.com" }
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
        { name: "Alankar Awadhiya", age: 27, email: "alanawadhiya007@gmail.com" },
        { name: "Atul", age: 25, email: "atulawadhiya@example.com" },
        { name: "Niharika", age: 35, email: "niharikaawadhiya@example.com" }
    ];
    
    const students = [
        { rollNo: 1, name: "Atul", grade: 85 },
        { rollNo: 2, name: "Ambrish", grade: 75 },
        { rollNo: 3, name: "Aniruddha", grade: 72 },
        { rollNo: 4, name: "Alice", grade: 70 }
        { rollNo: 5, name: "Aakash", grade: 65 }
        { rollNo: 6, name: "Anuraag", grade: 88 }
        { rollNo: 7, name: "Alankar", grade: 95 },
        { rollNo: 8, name: "Anmol", grade: 70 }
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

## License

Please consult the attached LICENSE file for details.  All rights not explicitly
granted by the Apache 2.0 License are reserved by the Original Author.