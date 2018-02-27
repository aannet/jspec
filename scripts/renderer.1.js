'use strict';

/* ----------------------   EXCEL ----------------------------- */
const excel = require("exceljs");
const docx = require("docx");
var fs = require('fs');

var workbook = new excel.Workbook();
// var filename = '../import/US-DatesCles.xlsx';
var filename = "C:\\AW\\webperso\\jspec\\import\\US-DatesCles.xlsx";
fs.access(filename, fs.constants.R_OK | fs.constants.W_OK, (err) => {
    console.log(err ? 'no access!' : 'can read/write');
});

const dataUsStatic = {
    labels : {
        usAsLabel : "En tant que2",
        usToLabel : "Afin de",
        usICanLabel : "Je peux",
        usCommentsLabel : "Commentaire"
    },
    usList : [
        {
            id : "ID_US_1",
            usAs : "pochtron",
            usTo : "de devenir saoul",
            usIcan : "boire cul sec",
            usComments : "Mes commentaires",
            rm: [
                {
                    id: "RM_01",
                    rmText: "Ma règle 1"
                },
                {
                    id: "RM_02",
                    rmText: "Ma règle 2"
                }
            ]
        },
        {
            id: "ID_US_2",
            usAs: "pochtron2",
            usTo: "de devenir saoul2",
            usIcan: "boire cul sec2",
            usComments: "Mes commentaires2",
            rm: [
                {
                    id: "RM_03",
                    rmText: "Ma règle 3"
                }
            ]
        }
    ]
};

const dataRMStatic = {
    rmList : [
        {
            id: "RM_01",
            rmText: "Ma règle 1"
        },
        {
            id: "RM_02",
            rmText: "Ma règle 2"
        },

    ]
};

let dataUSDynamic = new Object; //JSON des US
let dataRMDynamic = new Object; //JSON des RM

/**
 * ORDRE D'APPEL
 * loadUSSheet() > loadRMSheet() > generateDom()
 */
loadUSSheet();

function loadUSSheet() {

    workbook.xlsx.readFile(filename).then(function () {
        workbook.eachSheet(function (worksheet, sheetId) {

            console.log(worksheet.name);
            // console.log(sheetId);

            if (worksheet.name == "US") {
                var USSheet = workbook.getWorksheet('US');

                //WORK
                // var cellTest = USSheet.getCell('A2').value;
                // console.log(cellTest);

                let usListToAdd = new Array;

                USSheet.eachRow(function (row, rowNumber) {
                    // console.log('Row ' + rowNumber + ' = ' + JSON.stringify(row.values));

                    //set des colonnes en dur pour le moment
                    const COL_US_ID = 1;
                    const COL_US_AS = 2;
                    const COL_US_TO = 3;
                    const COL_US_ICAN = 4;
                    const COL_US_COMMENTS = 5;

                    //ENTETE
                    if (rowNumber == 1) {

                        dataUSDynamic.labels = {
                            usAsLabel: row.values[COL_US_AS],
                            usToLabel: row.values[COL_US_TO],
                            usICanLabel: row.values[COL_US_ICAN],
                            usCommentsLabel: row.values[COL_US_COMMENTS]
                        };
                    }
                    else {
                        const usToAdd = new Object;
                        if (row.values[COL_US_ID] !== undefined) {
                            usToAdd.id = row.values[COL_US_ID];
                        }
                        if (row.values[COL_US_AS] !== undefined) {
                            usToAdd.usAs = row.values[COL_US_AS];
                        }
                        if (row.values[COL_US_TO] !== undefined) {
                            usToAdd.usTo = row.values[COL_US_TO];
                        }
                        if (row.values[COL_US_ICAN] !== undefined) {
                            usToAdd.usIcan = row.values[COL_US_ICAN];
                        }
                        if (row.values[COL_US_COMMENTS] !== undefined) {
                            usToAdd.usComments = row.values[COL_US_COMMENTS];
                        }
                        usListToAdd.push(usToAdd);
                    }
                });

                dataUSDynamic.usList = usListToAdd;
                console.log('--dataUSDynamic--');
                console.log(dataUSDynamic);

                loadRMSheet();
            }
        });
    });
}


function loadRMSheet() {

    workbook.xlsx.readFile(filename).then(function () {
        workbook.eachSheet(function (worksheet, sheetId) {

            console.log(worksheet.name);
            // console.log(sheetId);

            if (worksheet.name == "RM") {
                var RMSheet = workbook.getWorksheet('RM');

                //WORK
                // var cellTest = RMSheet.getCell('A2').value;
                // console.log(cellTest);

                let rmListToAdd = new Array;

                RMSheet.eachRow(function (row, rowNumber) {
                    // console.log('Row ' + rowNumber + ' = ' + JSON.stringify(row.values));

                    //ENTETE
                    if (rowNumber == 1) {
                        //NO LABELS FOR NOW
                    }
                    else {
                        //set des colonnes en dur pour le moment
                        const COL_RM_ID = 3;
                        const COL_RM_TEXT = 4;

                        const rmToAdd = new Object;
                        if (row.values[COL_RM_ID] !== undefined) {
                            rmToAdd.id = row.values[COL_RM_ID];
                        }
                        if (row.values[COL_RM_TEXT] !== undefined) {
                            rmToAdd.usAs = row.values[COL_RM_TEXT];
                        }
                        rmListToAdd.push(rmToAdd);
                    }
                });

                dataRMDynamic.rmList = rmListToAdd;
                console.log('--dataRMDynamic--');
                console.log(dataRMDynamic);

                generateDom();
            }

            //finally we generate the DOM from the JSON we populated
             generateDom();
        });
    });

}



/* PARSE dataUsStatic to generate the DOM markup */
function generateDom() {
    // console.log("*** DOM ***");
    for (const i in dataUSDynamic.usList) {
        // console.log(dataUSDynamic.usList[i].usAs);

        //Init des labels
        const usAsLabel = dataUSDynamic.labels.usAsLabel; //EN TANT QUE 
        const usToLabel = dataUSDynamic.labels.usToLabel; //AFIN DE 
        const usICanLabel = dataUSDynamic.labels.usICanLabel; //JE PEUX
        const usCommentsLabel = dataUSDynamic.labels.usCommentsLabel; //COMMENTAIRE


        const USListElement = document.createElement("li");

        //AJOUT du <span>ID US</span>
        const span = document.createElement("span");
        span.appendChild(document.createTextNode(dataUSDynamic.usList[i].id));
        span.classList.add("id-us");
        USListElement.appendChild(span);

        //AJOUT du <ul class="inner-us"></ul>
        const usSubList = document.createElement("ul");
        usSubList.classList.add('inner-us');
        if (dataUSDynamic.usList[i].usAs !== undefined) {
            const li = document.createElement("li");
            const txt = `${usAsLabel} ${dataUSDynamic.usList[i].usAs}`;
            li.appendChild(document.createTextNode(txt)); //Ajout du texte au LI
            usSubList.appendChild(li);
        }
        if (dataUSDynamic.usList[i].usTo !== undefined) {
            const li = document.createElement("li");
            const txt = `${usToLabel} ${dataUSDynamic.usList[i].usTo}`;
            li.appendChild(document.createTextNode(txt)); //Ajout du texte au LI
            usSubList.appendChild(li);
        }
        if (dataUSDynamic.usList[i].usIcan !== undefined) {
            const li = document.createElement("li");
            const txt = `${usICanLabel} ${dataUSDynamic.usList[i].usIcan}`;
            li.appendChild(document.createTextNode(txt)); //Ajout du texte au LI
            usSubList.appendChild(li);
        }
        //COMMENTAIRES
        if (dataUSDynamic.usList[i].usComments !== undefined) {
            const li = document.createElement("li");
            li.classList.add("comments");
            const txt = `${usICanLabel} ${dataUSDynamic.usList[i].usComments}`;
            li.appendChild(document.createTextNode(txt)); //Ajout du texte au LI
            usSubList.appendChild(li);
        }

        USListElement.appendChild(usSubList); //Ajout du texte au LI
        
        //AJOUT du <ul></ul> général des US
        const USList = document.createElement("ul");
        USList.appendChild(USListElement); //Ajout du LI au UL

        //ON ajoute la liste créée au DOM
        document.getElementById('us-wrapper').appendChild(USList);

    }
}