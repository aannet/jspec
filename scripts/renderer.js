'use strict';


var _ = require('lodash');

/* ----------------------   EXCEL ----------------------------- */
const excel = require("exceljs");
const docx = require("docx");
var fs = require('fs');

var workbook = new excel.Workbook();
var filename = "C:\\AW\\webperso\\jspec\\import\\US-DatesCles.xlsx";
fs.access(filename, fs.constants.R_OK | fs.constants.W_OK, (err) => {
    console.log(err ? 'no access!' : 'can read/write');
});





let dataUSDynamic = new Object; //JSON des US
let dataRMDynamic = new Object; //JSON des RM


// ORDRE D'INCLUSION DES APPELS
// loadUSSheet() > loadRMSheet() > mergeRMWithUS() > generateDom()


loadUSSheet();


/**
 * loadUSSheet
 */
function loadUSSheet() {

    workbook.xlsx.readFile(filename).then(function () {

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
            
    });
}

/**
 * loadRMSheet
 */
function loadRMSheet() {

    workbook.xlsx.readFile(filename).then(function () {

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
                const COL_RM_IDUS = 2;
                const COL_RM_ID = 3;
                const COL_RM_TEXT = 4;

                const rmToAdd = new Object;

                if (row.values[COL_RM_IDUS] !== undefined) {
                    rmToAdd.idUS = row.values[COL_RM_IDUS];
                }
                if (row.values[COL_RM_ID] !== undefined) {
                    rmToAdd.id = row.values[COL_RM_ID];
                }
                if (row.values[COL_RM_TEXT] !== undefined) {
                    rmToAdd.rmText = row.values[COL_RM_TEXT];
                }
                rmListToAdd.push(rmToAdd);
            }
        });

        dataRMDynamic.rmList = rmListToAdd;
        console.log('--dataRMDynamic--');
        console.log(dataRMDynamic);

        // generateDom();
        mergeRMWithUS();


        // //finally we generate the DOM from the JSON we populated
        // generateDom();
    });

}

/**
 * Parse RM Json and inject each RM to the corresponding US entry in the US JSON 
 * 
 * use of lodash function find 
 * https://lodash.com/docs/4.17.4#find
 */
function mergeRMWithUS() {
    
    for (const i in dataRMDynamic.rmList) {
        let correspondingUS = _.find(dataUSDynamic.usList, ['id', dataRMDynamic.rmList[i].idUS]);

        //TODO rajouter un test isset correspondingUS
       
        // let rmToAdd = new Array;
        // dataRMDynamic.rmList.push(dataRMDynamic.rmList[i]);

        //si le tableau possède déjà une entrée on concatène le tableau sinon on créé la clé "correspondingUS.rmList"
        if (_.has(correspondingUS, 'rmList') === false ) {
            let rmToAdd = new Array;
            rmToAdd.push(dataRMDynamic.rmList[i]);
            correspondingUS.rmList = rmToAdd;
        }
        else {
            Array.prototype.push.call(correspondingUS.rmList, dataRMDynamic.rmList[i]);
        } 
    }

    //finally we generate the DOM from the JSON we populated
    generateDom();
}

/* PARSE dataUsStatic to generate the DOM markup */
/**
 * generateDom
 */
function generateDom() {

    //AJOUT du <ul class="all-us"></ul> général des US
    const USList = document.createElement("ul");
    USList.classList.add("all-us");

    for (const i in dataUSDynamic.usList) {
        // console.log(dataUSDynamic.usList[i].usAs);

        //Init des labels
        const usAsLabel = dataUSDynamic.labels.usAsLabel; //EN TANT QUE 
        const usToLabel = dataUSDynamic.labels.usToLabel; //AFIN DE 
        const usICanLabel = dataUSDynamic.labels.usICanLabel; //JE PEUX
        const usCommentsLabel = dataUSDynamic.labels.usCommentsLabel; //COMMENTAIRE


        const USListElement = document.createElement("li");
        USListElement.classList.add('wrapper-each-us');

        //AJOUT du <span>ID US</span>
        const span = document.createElement("span");
        span.appendChild(document.createTextNode(dataUSDynamic.usList[i].id));
        span.classList.add("id-us");
        USListElement.appendChild(span);

        //AJOUT du <ul class="inner-us"></ul>
        const usSubList = document.createElement("ul");
        usSubList.classList.add('inner-us');
            //EN TANT QUE
            if (dataUSDynamic.usList[i].usAs !== undefined) {
                const li = document.createElement("li");
                const txt = `${usAsLabel} ${dataUSDynamic.usList[i].usAs}`;
                li.appendChild(document.createTextNode(txt)); //Ajout du texte au LI
                li.classList.add('us-fragment');
                usSubList.appendChild(li);
            }
            //AFIN DE
            if (dataUSDynamic.usList[i].usTo !== undefined) {
                const li = document.createElement("li");
                const txt = `${usToLabel} ${dataUSDynamic.usList[i].usTo}`;
                li.appendChild(document.createTextNode(txt)); //Ajout du texte au LI
                li.classList.add('us-fragment');
                usSubList.appendChild(li);
            }
            //JE PEUX
            if (dataUSDynamic.usList[i].usIcan !== undefined) {
                const li = document.createElement("li");
                const txt = `${usICanLabel} ${dataUSDynamic.usList[i].usIcan}`;
                li.appendChild(document.createTextNode(txt)); //Ajout du texte au LI
                li.classList.add('us-fragment');
                usSubList.appendChild(li);
            }
            //COMMENTAIRES
            if (dataUSDynamic.usList[i].usComments !== undefined) {
                const li = document.createElement("li");
                li.classList.add("comments");
                const txt = `${usICanLabel} ${dataUSDynamic.usList[i].usComments}`;
                li.appendChild(document.createTextNode(txt)); //Ajout du texte au LI
                li.classList.add('us-fragment');
                usSubList.appendChild(li);
            }

        USListElement.appendChild(usSubList); //Ajout du sous UL au UL principal
        
        // **********************************************
        //                REGLES METIERS
        // **********************************************

        //AJOUT du <ul class="rm-list"></ul>
        const rmSubList = document.createElement("ul");
        rmSubList.classList.add('rm-list');

        if (dataUSDynamic.usList[i].rmList !== undefined) {
            for (const j in dataUSDynamic.usList[i].rmList) {
                const rmList = dataUSDynamic.usList[i].rmList[j]; //just shortcut
                const li = document.createElement("li");
                const txt = `${rmList.id} ${rmList.rmText}`;
                li.appendChild(document.createTextNode(txt)); //Ajout du texte au LI
                li.classList.add('wrapper-each-rm');
                rmSubList.appendChild(li);
            }

            USListElement.appendChild(rmSubList); 
        }


        USList.appendChild(USListElement); //Ajout du LI au UL

    }

    // **********************************************
    //               UL Global + body
    // **********************************************
    document.getElementById('us-wrapper').appendChild(USList);
}