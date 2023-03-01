var XLSX = require("xlsx");

class excelHandler {

    constructor(info){
        this.newWorkBook = XLSX.utils.book_new();
        if (!info.path) {
            this.path = "";
        } else {
            this.path = info.path;
        }
    }

    /* JSON TO EXCEL SHEET :: EXPORT SINGLE SHEET */
    async exportSingleSheet(excelConfig)
    {
        let resp;
        try {
            console.log(`excelHandler.exportSingleSheet() :: excelConfig: `, excelConfig);
            let {arrData, sheetName, arrFieldNames} = excelConfig;
            const newWB = this.newWorkBook;

            if(!arrData) arrData = [];
            if(!sheetName) sheetName = "Sheet1";
            if(!arrFieldNames) arrFieldNames = [];

            // const newObj = arrData.toObject();
            const txt = JSON.stringify(arrData);
            const newObj = JSON.parse(txt);
        
            /* Json to Excel package doesnt parse array strings, so handling separately by converting these array fields into string to avoid errors. */
            const newNewObj = [];
            newObj.forEach(function (val, ind) {
                // console.log(`excelHandler.exportSingleSheet() :: val: `, val, ` ind: `, ind);
                newNewObj.push(val);
                for (var key of Object.keys(val)) {
                    if (arrFieldNames.includes(key)) {
                        val[key] = val[key].join(', ');
                    }
                }
            });
            // console.log(`excelHandler.exportSingleSheet() :: newNewObj: `, newNewObj);

            var newWS = XLSX.utils.json_to_sheet(newObj);
            XLSX.utils.book_append_sheet(newWB, newWS, sheetName)//workbook name as param
        
            XLSX.writeFile(newWB, this.path)//file name as param
            resp = {
                error: false,
                message: `Data written to file`
            };
            console.log(`excelHandler.exportSingleSheet() :: resp: `, resp);
            return resp;
                
        } catch (error) {
            console.error(`excelHandler.exportSingleSheet() :: error: `, error);
            resp = {
                error: true,
                message: `Error : ${error}`
            };
            return resp;
            // throw error;
        }
    }

    /* JSON TO EXCEL SHEET :: EXPORT MULTIPLE SHEETS */
    async exportMultiSheets(excelConfig)
    {
        let resp;
        try {
            console.log(`excelHandler.exportMultiSheets() :: excelConfig: `, excelConfig);
            let {arrData, sheetName, arrFieldNames, isLastSheet} = excelConfig;
            const newWB = this.newWorkBook;

            if(!arrData) arrData = [];
            if(!sheetName) sheetName = "Sheet1";
            if(!arrFieldNames) arrFieldNames = [];
            if(!isLastSheet) isLastSheet = false;

            // const newObj = arrData.toObject();
            const txt = JSON.stringify(arrData);
            const newObj = JSON.parse(txt);
        
            /* Json to Excel package doesnt parse array strings, so handling separately by converting these array fields into string to avoid errors. */
            const newNewObj = [];
            newObj.forEach(function (val, ind) {
                // console.log(`excelHandler.exportMultiSheets() :: val: `, val, ` ind: `, ind);
                newNewObj.push(val);
                for (var key of Object.keys(val)) {
                    if (arrFieldNames.includes(key)) {
                        val[key] = val[key].join(', ');
                    }
                }
            });
            // console.log(`excelHandler.exportMultiSheets() :: newNewObj: `, newNewObj);

            var newWS = XLSX.utils.json_to_sheet(newObj);
            XLSX.utils.book_append_sheet(newWB, newWS, sheetName)//workbook name as param
            if (isLastSheet) {
                XLSX.writeFile(newWB, this.path)//file name as param
                resp = {
                    error: false,
                    message: `Data written to file`
                };
                console.log(`excelHandler.exportMultiSheets() :: resp: `, resp);
                return resp;
            }
        } catch (error) {
            console.error(`excelHandler.exportMultiSheets() :: error: `, error);
            resp = {
                error: true,
                message: `Error : ${error}`
            };
            return resp;
            // throw error;
        }
    }

    /* JSON TO EXCEL SHEET :: EXPORT COMMON */
    async exportSheets(excelConfig)
    {
        let resp;
        try {
            console.log(`excelHandler.exportSheets() :: excelConfig: `, excelConfig);
            let {arrData, sheetName, arrFieldNames, isLastSheet, isMultipleSheets} = excelConfig;
            const newWB = this.newWorkBook;

            if(!arrData) arrData = [];
            if(!sheetName) sheetName = "Sheet1";
            if(!arrFieldNames) arrFieldNames = [];
            if(!isLastSheet) isLastSheet = false;

            // const newObj = arrData.toObject();
            const txt = JSON.stringify(arrData);
            const newObj = JSON.parse(txt);
        
            /* Json to Excel package doesnt parse array strings, so handling separately by converting these array fields into string to avoid errors. */
            const newNewObj = [];
            newObj.forEach(function (val, ind) {
                // console.log(`excelHandler.exportSheets() :: val: `, val, ` ind: `, ind);
                newNewObj.push(val);
                for (var key of Object.keys(val)) {
                    if (arrFieldNames.includes(key)) {
                        val[key] = val[key].join(', ');
                    }
                }
            });
            // console.log(`excelHandler.exportSheets() :: newNewObj: `, newNewObj);

            var newWS = XLSX.utils.json_to_sheet(newObj);
            XLSX.utils.book_append_sheet(newWB, newWS, sheetName)//workbook name as param
            if (isMultipleSheets) {
                console.log(`excelHandler.exportSheets() :: isMultipleSheets: `, isMultipleSheets);
                if (isLastSheet) {
                    console.log(`excelHandler.exportSheets() :: isLastSheet: `, isLastSheet);
                    XLSX.writeFile(newWB, this.path)//file name as param
                    resp = {
                        error: false,
                        message: `Data written to file`
                    };
                    console.log(`excelHandler.exportSheets() :: resp: `, resp);
                    return resp;
                }
            }else{
                XLSX.writeFile(newWB, this.path)//file name as param
                resp = {
                    error: false,
                    message: `Data written to file`
                };
                console.log(`excelHandler.exportSheets() :: resp: `, resp);
                return resp;
            }
            
        } catch (error) {
            console.error(`excelHandler.exportSheets() :: error: `, error);
            resp = {
                error: true,
                message: `Error : ${error}`
            };
            return resp;
            // throw error;
        }
    }

    /* CONVERT BUFFER DATA TO JSON */
    bufferToJson (BUFFER) {
        let  resp;
        try {
            // console.log(`excelHandler.bufferToJson() :: BUFFER: `, BUFFER);
            /* Reading our test file */
            // var wb = XLSX.read(buffer);
        
            const wb = XLSX.readFile(this.path);
            // console.dir(`excelHandler.bufferToJson() :: Workbook(dir):`, wb);
            // console.log(`excelHandler.bufferToJson() :: Workbook: `, wb);
        
            let data = [];
            let JsonSheets = {};
        
            const sheets = wb.SheetNames;
            // console.log(`excelHandler.bufferToJson() :: sheets: `, sheets);
        
            for (let i = 0; i < sheets.length; i++) {
                data = [];
                const temp = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[i]]);
                // console.log(`excelHandler.bufferToJson() :: temp: `,
                //   temp,
                //   `wb.Sheets[${wb.SheetNames[i]}] : `,
                //   wb.Sheets[wb.SheetNames[i]]
                // );
            
                temp.forEach((res) => {
                    data.push(res);
                });
            
                JsonSheets[wb.SheetNames[i]] = data;
            }
        
            // Printing data
            // console.log(`excelHandler.bufferToJson() :: JsonData: `, data);
            // console.log(`excelHandler.bufferToJson() :: JsonSheets: `, JsonSheets);
            resp = {
                error: true,
                result: JsonSheets
            };
            return resp;            
        } catch (error) {
            console.error(`excelHandler.bufferToJson() :: error: `, error);
            resp = {
                error: true,
                message: `Error : ${error}`
            };
            return resp;
            // throw error;
        }
    }

}

module.exports = excelHandler;
