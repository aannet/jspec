'use strict';


/* ----------------------   EXCEL ----------------------------- */
const excel = require("exceljs");
const docx = require("docx");
var fs = require('fs');

var workbook2 = new excel.Workbook();
var filename2 = "C:\\AW\\webperso\\jspec\\import\\Tagerim-PlanTest-Lot1-V0.9-Recette.xlsx";
fs.access(filename2, fs.constants.R_OK | fs.constants.W_OK, (err) => {
    console.log(err ? 'no access!' : `Can read/write ${filename2}`);
});



workbook2.xlsx.readFile(filename2).then(function () {
    
    
    workbook2.eachSheet(function (worksheet, sheetId) {
        console.log(worksheet.name);
    });


    var USSheet = workbook2.getWorksheet("User Story");
    USSheet.eachRow(function (row, rowNumber) {
        console.log('Row ' + rowNumber + ' = ' + JSON.stringify(row.values));
        console.log('Row ' + rowNumber + ' = ' + row.values.length);
           
    });



    //lecture d'un sheet en particulier
    
});
