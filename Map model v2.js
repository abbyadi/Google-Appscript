function onOpen(){
  let ui = SpreadsheetApp.getUi();
  ui.createMenu('9% Sort').addItem('Nonprofit', 'nonProfitandAtRiskSort').addItem('Rural RHS & Native American', 'ruralSort').addItem('Rural Other','ruralOtherSort').addItem('Special Needs', 'splNeedsSort').addItem('Geographic Region','geoRegionSort').addSeparator().addItem('Reset Credit Counter','resetCreditCounter').addToUi();
};

let ss = SpreadsheetApp.getActiveSpreadsheet();
//Get Values of Counters in Array Form
let hTValues = ss.getRangeByName('HousingType_Counter').getValues();
let sAValues = ss.getRangeByName('SetAside_Counter').getValues();
let ruralHTValues = ss.getRangeByName('Rural_Counter').getValues();
let regionValues = ss.getRangeByName('GeoRegion_Counter').getValues();

//Get Credit Amounts in Map Form
let credit = {
  get setAside() {
    let mapSAsideArr = ss.getRangeByName('Set_Aside_Range').getValues().map(col => [col[0],col[7]]).slice(2,10);
    mapSAsideArr.splice(1,1);
    mapSAsideArr.splice(4,1);
    let subMapSAsideArr = new Map(mapSAsideArr);
    console.log(subMapSAsideArr);
    return subMapSAsideArr;
  },

  get region(){
    let mapGeoArr = ss.getRangeByName('Geo_Region_Range').getValues().map(col => [col[0],col[9]]).slice(1,13);
    let subMapGeoArr = new Map(mapGeoArr);
    console.log(subMapGeoArr);
    return subMapGeoArr;
  },

  get hsgType(){
    let appRndNum = ss.getSheetByName('Credit Amount').getRange(4,11).getValue();
    let hsgTypArr = new Array;
    if(appRndNum === 'R1'){
      hsgTypArr = ss.getRangeByName("Housing_Type_Range").getValues().map(col=>[col[0],col[5]]).slice(1);
    } else if(appRndNum === 'R2'){
      hsgTypArr = ss.getRangeByName("Housing_Type_Range").getValues().map(col=>[col[0],col[6]]).slice(1);
    }
    let subMapHsgTypArr = new Map(hsgTypArr);
    console.log(subMapHsgTypArr);
    return subMapHsgTypArr;
  },

  get ruralHsgTyp(){
    let appRndNum = ss.getSheetByName('Credit Amount').getRange(4,11).getValue();
    let ruralHsgTypArr = new Array;
    if(appRndNum === 'R1'){
      ruralHsgTypArr = ss.getRangeByName("Rural_Housing_Type_Range").getValues().map(col=>[col[0],col[5]]).slice(1);
    } else if(appRndNum === 'R2'){
      ruralHsgTypArr = ss.getRangeByName("Rural_Housing_Type_Range").getValues().map(col=>[col[0],col[6]]).slice(1);
    }
    let subMapRuralHsgTypArr = new Map(ruralHsgTypArr);
    console.log(subMapRuralHsgTypArr);
    return subMapRuralHsgTypArr;
  },
};
//Counter Values in Map Form
let counter = {
  get setAside() {
    let mapSAsideArr = ss.getRangeByName('SetAside_Counter').getValues().map(col => [col[1],col[2]]).slice(1);
    let subMapSAsideArr = new Map(mapSAsideArr);
    console.log(subMapSAsideArr);
    return subMapSAsideArr;
  },

  get hsgTyp() {
    let mapHsgTypArr = ss.getRangeByName('HousingType_Counter').getValues().map(col => [col[1],col[2]]).slice(1);
    let subMapHsgTypArr = new Map(mapHsgTypArr);
    console.log(subMapHsgTypArr);
    return subMapHsgTypArr;
  },

  get ruralHsgTyp() {
    let mapRuralHsgTypArr = ss.getRangeByName('Rural_Counter').getValues().map(col => [col[1],col[2]]).slice(1);
    let subMapRuralHsgTypArr = new Map(mapRuralHsgTypArr);
    console.log(subMapRuralHsgTypArr);
    return subMapRuralHsgTypArr;
  },

  get region() {
    let mapGeoRegArr = ss.getRangeByName('GeoRegion_Counter').getValues().map(col => [col[1],col[2]]).slice(1);
    let subMapGeoRegArr = new Map(mapGeoRegArr);
    console.log(subMapGeoRegArr);
    return subMapGeoRegArr;
  },
};
//Declare Counters
let hsgTypCounter = counter.hsgTyp;
let sACounter = counter.setAside;
let ruralCounter = counter.ruralHsgTyp;
let regionCounter = counter.region;

// Function to get raw data Values in Array Form
function rawDataWrapper(){
  let rawDataSheet = ss.getSheetByName('Raw Data');
  let rawDataRange = rawDataSheet.getDataRange();
  let rawDataRangeValues = rawDataRange.getValues();
  return rawDataRangeValues;
}

function setSheetandHeader(sheetName){
  ss.insertSheet(sheetName,ss.getNumSheets()+1);
  let inputSheet = ss.getSheetByName(sheetName);
  let rawDataSheetValues = rawDataWrapper();
  let setAsideCredits = credit.setAside;
  let geoRegionCredits = credit.region;
  if(sheetName == 'Nonprofit' || sheetName == 'Rural' || sheetName == 'At-Risk' || sheetName == 'Special Needs'){
    if(sheetName == 'Rural'){
      let ruralCredit = setAsideCredits.get('RHS & HOME Apportionment') + setAsideCredits.get('Native American')+setAsideCredits.get('Other');
      inputSheet.getRange(1,1).setValue(ruralCredit).setHorizontalAlignment('left').setNumberFormat('#,###');
    } else{
      inputSheet.getRange(1,1).setValue(setAsideCredits.get(sheetName)).setHorizontalAlignment('left').setNumberFormat('#,###');
    }
  } else {
    inputSheet.getRange(1,1).setValue(geoRegionCredits.get(sheetName)).setHorizontalAlignment('left').setNumberFormat('#,###');
    //Fund up to 125% of credit amount available.
    inputSheet.getRange(2,1).setValue(Math.round(geoRegionCredits.get(sheetName)*1.25)).setHorizontalAlignment('left').setNumberFormat('#,###');
    inputSheet.getRange(3,1).setValue(Math.round(geoRegionCredits.get(sheetName)*1.25)).setHorizontalAlignment('left').setNumberFormat('#,###');
    inputSheet.getRange(4,1).setValue(0).setHorizontalAlignment('left').setNumberFormat('0');
    inputSheet.getRange(2,2).setValue('x 125%').setHorizontalAlignment('right');
    inputSheet.getRange(3,2).setValue('Balance Remaining').setHorizontalAlignment('right');
    inputSheet.getRange(4,2).setValue('Projects Funded via Geo Credits').setHorizontalAlignment('right');
  };
  inputSheet.getRange(1,2).setValue(sheetName).setHorizontalAlignment('right');
  inputSheet.getRange(inputSheet.getLastRow()+1,1,1,rawDataSheetValues[0].length).setValues(rawDataSheetValues.slice(0,1)).setFontWeight('bold').setHorizontalAlignment('center').setWrap(true);
  inputSheet.getRange(inputSheet.getLastRow(), inputSheet.getLastColumn()+1).setValue('FUND/SKIP').setFontWeight('bold').setHorizontalAlignment('center').setWrap(true);
}

function getSetSheetData(sheetName){
  let inputSheet = ss.getSheetByName(sheetName);
  let rawDataValues = rawDataWrapper();
  let regEx = new RegExp(sheetName,'i')
  let arrRawDataToInput = new Array;
  if(sheetName == 'Nonprofit' || sheetName == 'Rural' || sheetName == 'At-Risk' || sheetName == 'Special Needs'){
    for(let j=0; j<rawDataValues.length; j++){
      if(regEx.test(rawDataValues[j][7])){
        arrRawDataToInput.push(rawDataValues[j]);
      } 
    };
  } else {
    for(let j=0; j<rawDataValues.length; j++){
      if(regEx.test(rawDataValues[j][8])){
        arrRawDataToInput.push(rawDataValues[j]);
      } 
    };
  }
  if(arrRawDataToInput.length === 0){
    console.log('No projects applied in Region')
    let textFinder = inputSheet.createTextFinder('TOTAL STATE CREDIT REQUESTED');
    let foundTxtA1 = textFinder.findNext().getA1Notation();
    let foundTxtRange = inputSheet.getRange(foundTxtA1);
    let foundTxtRow = foundTxtRange.getRow();
    let foundTxtCol = foundTxtRange.getColumn();
    inputSheet.insertColumnsAfter(foundTxtCol,1);
    inputSheet.getRange(foundTxtRow,foundTxtCol+1).setValue('COMBINED CREDIT REQUESTED').setFontWeight('bold').setHorizontalAlignment('center').setWrap(true);
  } else {
    inputSheet.getRange(inputSheet.getLastRow()+1,1,arrRawDataToInput.length,arrRawDataToInput[0].length).setValues(arrRawDataToInput).setFontWeight('normal').setHorizontalAlignment('right').setWrap(false).setNumberFormat('#,###');
    //Create 'Combined Credit Requested' Column
    let textFinder = inputSheet.createTextFinder('TOTAL STATE CREDIT REQUESTED');
    let foundTxtA1 = textFinder.findNext().getA1Notation();
    let foundTxtRange = inputSheet.getRange(foundTxtA1);
    let foundTxtRow = foundTxtRange.getRow();
    let foundTxtCol = foundTxtRange.getColumn();
    inputSheet.insertColumnsAfter(foundTxtCol,1);
    inputSheet.getRange(foundTxtRow,foundTxtCol+1).setValue('COMBINED CREDIT REQUESTED').setFontWeight('bold').setHorizontalAlignment('center').setWrap(true);
    inputSheet.getRange(foundTxtRow+1,foundTxtCol+1,inputSheet.getLastRow()-foundTxtRow,1).setFormulaR1C1("=((R[0]C[-2]*10)+R[0]C[-1])/10");
    //Set Formatting
    inputSheet.setColumnWidths(1,9,100);
    inputSheet.autoResizeColumn(2).autoResizeColumns(9,4).hideColumns(14,inputSheet.getLastColumn()-14);
    inputSheet.hideColumns(11,2);
    inputSheet.getRange('H:H').setNumberFormat('##.##%');
  }
};

function fundSA(sheetName){
  let inputSheet = ss.getSheetByName(sheetName);
  let x = inputSheet.getRange(1,1).getValue();
  let dataRange = inputSheet.getDataRange();

  //Sort first and then get Value
  if(sheetName == 'Nonprofit'){
    dataRange.offset(2,0).sort([{column: 7, ascending: false},{column: 9, ascending: true}, {column: 8, ascending: false}]);
  } else{
    dataRange.offset(2,0).sort([{column: 7, ascending: false}, {column: 8, ascending: false}]);
  }
  let dataRangeValues = dataRange.getValues();
  
  //Funding Code
  for (let k=2; k<dataRangeValues.length; k++){
    if (x >= 1){
      inputSheet.getRange(k+1,inputSheet.getLastColumn()).setValue('Fund').setHorizontalAlignment('normal');
      inputSheet.getRange(k+1,1,1,inputSheet.getLastColumn()).setBackground('#9bbb59');
      if (k===2) {
        ss.getSheetByName('Funded Projects').appendRow(['', sheetName]);
      }
      ss.getSheetByName('Funded Projects').appendRow(dataRangeValues[k]);
      x = x - dataRangeValues[k][3];
      console.log(x);
      if(dataRangeValues[k][2] === 'Large Family' && dataRangeValues[k][14] === 'Yes'){
        let lrgFamHgOppValue = hsgTypCounter.get('Large Family High Opportunity');
        let lrgFamValue = hsgTypCounter.get('Large Family');
        hsgTypCounter.set('Large Family High Opportunity', lrgFamHgOppValue-dataRangeValues[k][5]);
        hsgTypCounter.set('Large Family', lrgFamValue-dataRangeValues[k][5]);
        console.log(hsgTypCounter.get('Large Family High Opportunity'));
      } else {
        for (const bit of hsgTypCounter.entries()){
          if(bit[0] == dataRangeValues[k][2]){
            hsgTypCounter.set(bit[0], bit[1]-dataRangeValues[k][5]);
            console.log(hsgTypCounter.get(bit[0]));
          }
        }
      }
    } else {
      inputSheet.getRange(k+1,inputSheet.getLastColumn()).setValue('Skip').setHorizontalAlignment('normal');
    }
  }
  //Set Counters
  for(let i=0; i<sAValues.length; i++){
    if(sAValues[i][1] == sheetName){
      ss.getSheetByName('CreditCounter').getRange(i+1,3).setValue(x).setNumberFormat('#,###');
    }
  }
  for(let j=1; j<hTValues.length; j++){
    hTValues[j][2] = hsgTypCounter.get(hTValues[j][1]);
  }
  ss.getRangeByName('HousingType_Counter').setValues(hTValues).setNumberFormat('#,###');
};

function nonProfitandAtRiskSort(){
  let arrSANames = ['Nonprofit','At-Risk'];
  let hsgTypCounter = counter.hsgTyp;
  for(const element of arrSANames){
    setSheetandHeader(element);
    getSetSheetData(element);
    let inputSheet = ss.getSheetByName(element);
    fundSA(element);
  }
};

function splNeedsSort(){
  setSheetandHeader('Special Needs');
  getSetSheetData('Special Needs');
  let inputSheet = ss.getSheetByName('Special Needs');
  let regExSplNeeds = new RegExp('Special','i');
  let nonProfitSheet = ss.getSheetByName('Nonprofit');
  let splNeedsArr = new Array;
  let nonProfitDataValues = nonProfitSheet.getRange(3,1,nonProfitSheet.getLastRow()-2,nonProfitSheet.getLastColumn()).getValues();
  for(let z=0; z<nonProfitDataValues.length; z++){
    if(regExSplNeeds.test(nonProfitDataValues[z][2]) && nonProfitDataValues[z][26] !=='Fund'){
      splNeedsArr.push(nonProfitDataValues[z]);
    }
  }
  if(splNeedsArr.length>=1){
    inputSheet.getRange(inputSheet.getLastRow()+1,1,splNeedsArr.length,splNeedsArr[0].length).setValues(splNeedsArr).setFontWeight('normal').setHorizontalAlignment('right').setWrap(false).setNumberFormat('#,###');
    inputSheet.getRange('H:H').setNumberFormat('##.##%');
  }
  fundSA('Special Needs');
};

/**Function to sort Rural Set-Aside for "Native-American" and "RHS & HOME" Projects  **/
function ruralSort(){
  setSheetandHeader('Rural');
  getSetSheetData('Rural');
  ss.getSheetByName('Funded Projects').appendRow(['', 'Rural']);
  let inputSheet = ss.getSheetByName('Rural');
  let dataRange = inputSheet.getDataRange();
  inputSheet.unhideColumn(inputSheet.getRange('O1:O'));
  dataRange.offset(2,0).sort([{column: 7, ascending: false},{column: 8, ascending: false}]);
  let dataRangeValues = dataRange.getValues();
  //Fund RHS HOME Apportionment
  let sectionRegEx = new RegExp('Section','i');
  let homeRegex = new RegExp('HOME');
  let rhsAmValues = dataRangeValues.filter(item => sectionRegEx.test(item[8]) || homeRegex.test(item[8])); //Filtering for RHS and HOME  projects
  rhsAmValues.forEach((row, i, arr) => {
    if (sACounter.get('RHS & HOME Apportionment') >= 1 && sACounter.get('Rural') >= 1) { //Checking if there is more than $1 in RHS & HOME credit counter and in Rural Counter.
      if (chkRrlHTisNeg(row) === false) { //false means there is more than $1 and Hsg type is not negative
        chkRrlHTAndSetCounters(row);
        inputSheet.createTextFinder(row[0]).findNext().offset(0, 26).setValue('Fund').setHorizontalAlignment('normal').getDataRegion(SpreadsheetApp.Dimension.COLUMNS).setBackground('#9bbb59');
        ss.getSheetByName('Funded Projects').appendRow(row);
        let x = sACounter.get('RHS & HOME Apportionment');
        let y = sACounter.get('Rural');
        sACounter.set('RHS & HOME Apportionment', x - row[3]);
        sACounter.set('Rural', y - row[3]);
      } else if (chkRrlHTisNeg(row) === true) { //true means there is less than $1 and Hsg type is negative
        let count = i+1;
        if(count === arr.length) { // if the row being evaluated is the last project in the filtered set and funds if true.
          chkRrlHTAndSetCounters(row);
          inputSheet.createTextFinder(row[0]).findNext().offset(0, 26).setValue('Fund').setHorizontalAlignment('normal').offset(0,-26).getDataRegion(SpreadsheetApp.Dimension.COLUMNS).setBackground('#9bbb59');
          ss.getSheetByName('Funded Projects').appendRow(row);
          let x = sACounter.get('RHS & HOME Apportionment');
          let y = sACounter.get('Rural');
          sACounter.set('RHS & HOME Apportionment', x - row[3]);
          sACounter.set('Rural', y - row[3]);
        } else {
          chkRestofData(row, arr, i, inputSheet, 'RHS & HOME Apportionment'); //checks rest of data set and inputs "skip HT" if it finds project with a different Hsg type and same or higher score.
        }
      }
    }
  })
  let nativeAmRegEx = new RegExp('Native American','i');
  let nativeAmValues = dataRangeValues.filter(item => nativeAmRegEx.test(item[8])); //Filtering for Native American  projects
  nativeAmValues.forEach((row, i, arr) => { 
    if (sACounter.get('Native American') >= 1 && sACounter.get('Rural') >= 1) { //Checking if there is more than $1 in Native American credit counter and in Rural Counter.
      if (chkRrlHTisNeg(row) === false) { //false means there is more than $1 and Hsg type is not negative
        chkRrlHTAndSetCounters(row);
        inputSheet.createTextFinder(row[0]).findNext().offset(0, 26).setValue('Fund').setHorizontalAlignment('normal').getDataRegion(SpreadsheetApp.Dimension.COLUMNS).setBackground('#9bbb59');
        ss.getSheetByName('Funded Projects').appendRow(row);
        let x = sACounter.get('Native American');
        let y = sACounter.get('Rural');
        sACounter.set('Native American', x - row[3]);
        sACounter.set('Rural', y - row[3]);
      } else if (chkRrlHTisNeg(row) === true) { //true means there is less than $1 and Hsg type is negative
        let count = i+1;
        if(count === arr.length) { // if the row being evaluated is the last project in the filtered set and funds if true
          chkRrlHTAndSetCounters(row);
          inputSheet.createTextFinder(row[0]).findNext().offset(0, 26).setValue('Fund').setHorizontalAlignment('normal').getDataRegion(SpreadsheetApp.Dimension.COLUMNS).setBackground('#9bbb59');
          ss.getSheetByName('Funded Projects').appendRow(row);
          let x = sACounter.get('Native American');
          let y = sACounter.get('Rural');
          sACounter.set('Native American', x - row[3]);
          sACounter.set('Rural', y - row[3]);
        } else {
          chkRestofData(row, arr, i, inputSheet, 'Native American') //checks rest of data set and inputs "skip HT" if it finds project with a different Hsg type and same or higher score.
        }
      }
    }
  })
  setCounterValuesToSheet();
};

//Function to Sort and Fund Rural Projects after RHS and Native American projects are funded
function ruralOtherSort(){
  let inputSheet = ss.getSheetByName('Rural');
  let dataRange = inputSheet.getDataRange().offset(2,0);
  let dataRangeValues = dataRange.getValues();

  let otherValues = dataRangeValues.filter(item => item[26] != 'Fund');
  otherValues.forEach((row, i, arr) => {
    if (sACounter.get('Rural') >= 1) { //Checking if there is more than $1 in Rural Counter.
      if (chkRrlHTisNeg(row) === false) { //false means there is more than $1 and Hsg type is not negative
        chkRrlHTAndSetCounters(row);
        inputSheet.createTextFinder(row[0]).findNext().offset(0, 26).setValue('Fund').setHorizontalAlignment('normal').getDataRegion(SpreadsheetApp.Dimension.COLUMNS).setBackground('#9bbb59');
        ss.getSheetByName('Funded Projects').appendRow(row);
        let x = sACounter.get('Rural');
        sACounter.set('Rural', x - row[3]);
      } else if (chkRrlHTisNeg(row) === true) { //true means there is less than $1 and Hsg type is negative
        let count = i+1;
        if(count === arr.length) { //Checking if the chkRrlHTisNeg function has reached the last item in the data set and funds if true
          chkRrlHTAndSetCounters(row);
          inputSheet.createTextFinder(row[0]).findNext().offset(0, 26).setValue('Fund').setHorizontalAlignment('normal').getDataRegion(SpreadsheetApp.Dimension.COLUMNS).setBackground('#9bbb59');
          ss.getSheetByName('Funded Projects').appendRow(row);
          let x = sACounter.get('Rural');
          sACounter.set('Rural', x - row[3]);
        } else {
          chkRestofData(row, arr, i, inputSheet, 'Rural') //checks rest of data set and inputs "skip HT" if it finds project with a different Hsg type and same or higher score.
        }
      }
    }
  })
  setCounterValuesToSheet();
};

/**Function to check if the counters in rural housing type is less than or equal to 1. Mark as Skip/HT if so.*/
function chkRrlHTisNeg(dataRangeValue){
  let inputSheet = ss.getSheetByName('Rural');
  if(dataRangeValue[2] == 'Large Family' && dataRangeValue[14] == 'Yes'){
    if(ruralCounter.get('Large Family High Opportunity')<= 1){
      //inputSheet.getRange(i+1,inputSheet.getLastColumn()).setValue('Skip H/T').setHorizontalAlignment('normal');
      return true;
    } else{
      return false;
    }
  } else if(dataRangeValue[12] == 'Acquisition/Rehabilitation' || dataRangeValue[12] == 'Rehabilitation-Only'){
    if(ruralCounter.get('Acquisition/Rehabilitation')<= 1){
      //inputSheet.getRange(i+1,inputSheet.getLastColumn()).setValue('Skip H/T').setHorizontalAlignment('normal');
      return true
    } else {
      return false;
    }
  } else if(dataRangeValue[2] == 'Seniors'){
    if(ruralCounter.get('Seniors')<= 1){
      //inputSheet.getRange(i+1,inputSheet.getLastColumn()).setValue('Skip H/T').setHorizontalAlignment('normal');
      return true
    } else {
      return false;
    }
  } else {
    return false;
  }
};

/**Function to check the Housing type in Rural Tab and deduct the amount from Credit Counter Map if function chkRrlHtisNeg does not return true;*/
function chkRrlHTAndSetCounters(dataRangeValue){
  if(dataRangeValue[2] === 'Large Family' && dataRangeValue[14] === 'Yes'){
    setCounters('hsgTyp','Large Family High Opportunity',dataRangeValue[5]);
    setCounters('hsgTyp','Large Family',dataRangeValue[5])
    setCounters('ruralHsgTyp','Large Family High Opportunity',dataRangeValue[3]);
  } else if(dataRangeValue[12] == 'Acquisition/Rehabilitation' || dataRangeValue[12] == 'Rehabilitation-Only') {
    setCounters('ruralHsgTyp','Acquisition/Rehabilitation',dataRangeValue[3]);
    setCounters('hsgTyp',dataRangeValue[2],dataRangeValue[5]);
    setCounters('ruralHsgTyp',dataRangeValue[2],dataRangeValue[3])
  } else {
    setCounters('hsgTyp',dataRangeValue[2],dataRangeValue[5]);
    setCounters('ruralHsgTyp',dataRangeValue[2],dataRangeValue[3]);
  }
};
/** Sub Function to be used in the Rural Sort to check rest of the rows. It evaluates if any following data row has point score which is at least equal to current data row and then it checks if the housing type is different. If true it posts 'Skip H/T' in the Fund column. 
 * @param {((string|number)[]} row - current data row that is being evaluated
 * @param {((string|number)[][])} arr - array of data passed to function
 * @param {number} index - index position of the current data row that is being evaluated in the array
 * @param {Object} sheet - input sheet into which data will be entered
 * @param {string} sAName - name of the Rural set-aside
*/
function chkRestofData(row, arr, index, sheet, sAName) {
  let count = index+1
  while (count < arr.length - 1) {
    if (arr[count][6] >= row[6]) {
      if (row[2] === "Seniors") {
        if (arr[count][2] !== 'Seniors') {
          sheet.createTextFinder(row[0]).findNext().offset(0, 26).setValue('Skip H/T').setHorizontalAlignment('normal');
          return
        }
      } else if (row[2] === 'Large Family' && row[14] === 'Yes') {
        if (arr[count][2] !== 'Large Family' && arr[count][14] !== 'Yes') {
          sheet.createTextFinder(row[0]).findNext().offset(0, 26).setValue('Skip H/T').setHorizontalAlignment('normal');
          return
        }
      } else if (row[12] == 'Acquisition/Rehabilitation' || row[12] == 'Rehabilitation-Only') {
        if (arr[count][12] == 'Acquisition/Rehabilitation' || arr[count][12] == 'Rehabilitation-Only') {
          sheet.createTextFinder(row[0]).findNext().offset(0, 26).setValue('Skip H/T').setHorizontalAlignment('normal');
          return
        }
      }
    }
    count++
  }
  chkRrlHTAndSetCounters(row);
  sheet.createTextFinder(row[0]).findNext().offset(0, 26).setValue('Fund').setHorizontalAlignment('normal').getDataRegion(SpreadsheetApp.Dimension.COLUMNS).setBackground('#9bbb59');
  if(sAName = 'Rural'){
    ss.getSheetByName('Funded Projects').appendRow(row);
    let x = sACounter.get('Rural');
    sACounter.set('Rural', x - row[3]);
  } else {
    ss.getSheetByName('Funded Projects').appendRow(row);
    let x = sACounter.get(sAName);
    let y = sACounter.get('Rural');
    sACounter.set(sAName, x - row[3]);
    sACounter.set('Rural', y - row[3]);
  }
}

function geoRegionSort(){
  ss.getSheetByName('Funded Projects').appendRow(['', 'Geographic Region']);
  let appRndNum = ss.getSheetByName('Credit Amount').getRange(4,11).getValue();
  let geoRegArr = Array.from(regionCounter.keys());
  let fundDirArr;
  if(appRndNum === 'R1'){
    fundDirArr = geoRegArr.reverse();
  } else{
    fundDirArr = geoRegArr
  }
  //let geoRegArr = ['San Francisco County','Northern Region', 'East Bay Region'];
  for(const element of fundDirArr){ /**loop to go through each Region*/
    setSheetandHeader(element);
    getSetSheetData(element);
    let inputSheet = ss.getSheetByName(element);
    let fundedData = ss.getSheetByName('Funded Projects').getDataRange().getValues();
    let dataRange = inputSheet.getDataRange();
    dataRange.offset(5,0).sort([{column: 7, ascending: false}, {column: 8, ascending: false}]);
    let dataRangeValues = dataRange.getValues();
    let creditAmt125Pct = dataRangeValues[1][0];
    let balanceAmt = dataRangeValues[2][0];
    let numDealFunded = dataRangeValues[3][0];
    for(let i=5; i<dataRangeValues.length; i++){ /**loop to set Funded projects under Set-Aside*/
      for(j=0; j<fundedData.length; j++){
        if(dataRangeValues[i][0] == fundedData[j][0]){
          inputSheet.getRange(i+1,inputSheet.getLastColumn()).setValue('Fund S/A').setHorizontalAlignment('normal');
        }
      }
    }
    /**First Round of funding deals. Fund housing type even if negative if highest tiebreaker and first to get funded*/
    dataRange = inputSheet.getDataRange();
    dataRangeValues = dataRange.getValues();
    //console.log(dataRangeValues);
    if(dataRangeValues.length >= 6){ /**checks if there are any projects applying in the region*/
      if(dataRangeValues[3][0] == 0){ /**checks if this is the first project in the region */
        for(let i=5; i<dataRangeValues.length; i++){ /**loop to go through each project in region*/
          if(i == 5 && !dataRangeValues[i][26].includes('Fund')){
            if(dataRangeValues[i][5]<=balanceAmt){
              chkHTAndSetCounters(dataRangeValues[i]);
              inputSheet.getRange(i+1,inputSheet.getLastColumn()).setValue('Fund').setHorizontalAlignment('normal');
              inputSheet.getRange(i+1,1,1,inputSheet.getLastColumn()).setBackground('#9bbb59');
              ss.getSheetByName('Funded Projects').appendRow(dataRangeValues[i]);
              balanceAmt -= dataRangeValues[i][5];
              numDealFunded++;
              inputSheet.getRange(3,1).setValue(balanceAmt)
              inputSheet.getRange(4,1).setValue(numDealFunded);
              setCounterValuesToSheet();
              dataRange = inputSheet.getDataRange();
              dataRangeValues = dataRange.getValues();
              break;
            } else{
              inputSheet.getRange(i+1,inputSheet.getLastColumn()).setValue('Skip 125').setHorizontalAlignment('normal');
              dataRange = inputSheet.getDataRange();
              dataRangeValues = dataRange.getValues();
            }
          }else if(!dataRangeValues[i][26].includes('Fund')){
            if(dataRangeValues[i-1][26].includes('Skip 125')){ /**checks if previous project that got skipped has TB greater than 75% of current project being evaluated*/
              let skipDealTB = dataRangeValues[i-1][7];
              let currDealTB = dataRangeValues[i][7];
              if(currDealTB >= 0.75*skipDealTB){
                if(dataRangeValues[i][5]<=balanceAmt){ 
                  if(chkHTisNeg(element, dataRangeValues[i], i, dataRangeValues) === false){
                    chkHTAndSetCounters(dataRangeValues[i]);
                    inputSheet.getRange(i+1,inputSheet.getLastColumn()).setValue('Fund').setHorizontalAlignment('normal');
                    inputSheet.getRange(i+1,1,1,inputSheet.getLastColumn()).setBackground('#9bbb59');
                    ss.getSheetByName('Funded Projects').appendRow(dataRangeValues[i]);
                    balanceAmt -= dataRangeValues[i][5];
                    numDealFunded++
                    inputSheet.getRange(3,1).setValue(balanceAmt)
                    inputSheet.getRange(4,1).setValue(numDealFunded);
                    setCounterValuesToSheet();
                    dataRange = inputSheet.getDataRange();
                    dataRangeValues = dataRange.getValues();
                    break;
                  } else if(chkHTisNeg(element,dataRangeValues[i], i, dataRangeValues) === true) {
                    inputSheet.getRange(i+1,inputSheet.getLastColumn()).setValue('Skip H/T').setHorizontalAlignment('normal');
                    dataRange = inputSheet.getDataRange();
                    dataRangeValues = dataRange.getValues();
                  }
                }else{
                  inputSheet.getRange(i+1,inputSheet.getLastColumn()).setValue('Skip 125').setHorizontalAlignment('normal');
                  dataRange = inputSheet.getDataRange();
                  dataRangeValues = dataRange.getValues();
                }
              }else{
                inputSheet.getRange(i+1,inputSheet.getLastColumn()).setValue('Skip 75%TB').setHorizontalAlignment('normal');
                dataRange = inputSheet.getDataRange();
                dataRangeValues = dataRange.getValues();
              }
            } else {
              if(dataRangeValues[i][5]<=balanceAmt){
                if(chkHTisNeg(element, dataRangeValues[i], i, dataRangeValues) === false){
                  chkHTAndSetCounters(dataRangeValues[i]);
                  inputSheet.getRange(i+1,inputSheet.getLastColumn()).setValue('Fund').setHorizontalAlignment('normal');
                  inputSheet.getRange(i+1,1,1,inputSheet.getLastColumn()).setBackground('#9bbb59');
                  ss.getSheetByName('Funded Projects').appendRow(dataRangeValues[i]);
                  balanceAmt -= dataRangeValues[i][5];
                  numDealFunded++
                  inputSheet.getRange(3,1).setValue(balanceAmt)
                  inputSheet.getRange(4,1).setValue(numDealFunded);
                  setCounterValuesToSheet();
                  dataRange = inputSheet.getDataRange();
                  dataRangeValues = dataRange.getValues();
                  break;
                } else if(chkHTisNeg(element,dataRangeValues[i], i, dataRangeValues) === true) {
                  inputSheet.getRange(i+1,inputSheet.getLastColumn()).setValue('Skip H/T').setHorizontalAlignment('normal');
                  dataRange = inputSheet.getDataRange();
                  dataRangeValues = dataRange.getValues();
                }
              }else{
                inputSheet.getRange(i+1,inputSheet.getLastColumn()).setValue('Skip 125').setHorizontalAlignment('normal');
                dataRange = inputSheet.getDataRange();
                dataRangeValues = dataRange.getValues();
              }
            }
          }
        }
      }
    }
  }
  /**Second to Fifth go through to fund deals if balance remaining*/
  let runCycle = 1;
  while(runCycle<=4){
    for(const element of fundDirArr){ /**loop to go through each Region*/
      let inputSheet = ss.getSheetByName(element);
      let dataRange = inputSheet.getDataRange();
      let dataRangeValues = dataRange.getValues();
      let creditAmt125Pct = dataRangeValues[1][0];
      let balanceAmt = dataRangeValues[2][0];
      let numDealFunded = dataRangeValues[3][0];
      if(dataRangeValues.length >= 6){ /**checks if there are any projects applying in the region*/
        if(numDealFunded === runCycle){
          for(let i=5; i<dataRangeValues.length; i++){ /**loop to go through each project in region*/
            if(dataRangeValues[i][26] === ''){
              if(dataRangeValues[i-1][26].includes('Skip 125')){ /**checks if previous project that got skipped has TB greater than 75% of current project being evaluated*/
                let skipDealTB = dataRangeValues[i-1][7];
                let currDealTB = dataRangeValues[i][7];
                if(currDealTB >= 0.75*skipDealTB){
                  if(dataRangeValues[i][5]<=balanceAmt){
                    if(regionCounter.get(element) > 0) { /**checks if region has run out of Credits.*/
                      if(chkHTisNeg(element, dataRangeValues[i], i, dataRangeValues) === false){
                        chkHTAndSetCounters(dataRangeValues[i]);
                        inputSheet.getRange(i+1,inputSheet.getLastColumn()).setValue('Fund').setHorizontalAlignment('normal');
                        inputSheet.getRange(i+1,1,1,inputSheet.getLastColumn()).setBackground('#9bbb59');
                        ss.getSheetByName('Funded Projects').appendRow(dataRangeValues[i]);
                        balanceAmt -= dataRangeValues[i][5];
                        numDealFunded++
                        inputSheet.getRange(3,1).setValue(balanceAmt)
                        inputSheet.getRange(4,1).setValue(numDealFunded);
                        setCounterValuesToSheet();
                        dataRange = inputSheet.getDataRange();
                        dataRangeValues = dataRange.getValues();
                        break;
                      } else if(chkHTisNeg(element,dataRangeValues[i], i, dataRangeValues) === true) {
                        inputSheet.getRange(i+1,inputSheet.getLastColumn()).setValue('Skip H/T').setHorizontalAlignment('normal');
                        dataRange = inputSheet.getDataRange();
                        dataRangeValues = dataRange.getValues();
                      }
                    } else {
                      inputSheet.getRange(i+1,inputSheet.getLastColumn()).setValue('Skip Negative').setHorizontalAlignment('normal');
                      dataRange = inputSheet.getDataRange();
                      dataRangeValues = dataRange.getValues();
                      break;
                    }
                  }else{
                    inputSheet.getRange(i+1,inputSheet.getLastColumn()).setValue('Skip 125').setHorizontalAlignment('normal');
                    dataRange = inputSheet.getDataRange();
                    dataRangeValues = dataRange.getValues();
                  }
                }else{
                  inputSheet.getRange(i+1,inputSheet.getLastColumn()).setValue('Skip 75%TB').setHorizontalAlignment('normal');
                  dataRange = inputSheet.getDataRange();
                  dataRangeValues = dataRange.getValues();
                }
              } else {
                if(dataRangeValues[i][5]<=balanceAmt){
                  if(regionCounter.get(element) > 0) { /**checks if region has run out of Credits.*/
                    if(chkHTisNeg(element, dataRangeValues[i], i, dataRangeValues) === false){
                      chkHTAndSetCounters(dataRangeValues[i]);
                      inputSheet.getRange(i+1,inputSheet.getLastColumn()).setValue('Fund').setHorizontalAlignment('normal');
                      inputSheet.getRange(i+1,1,1,inputSheet.getLastColumn()).setBackground('#9bbb59');
                      ss.getSheetByName('Funded Projects').appendRow(dataRangeValues[i]);
                      balanceAmt -= dataRangeValues[i][5];
                      numDealFunded++
                      inputSheet.getRange(3,1).setValue(balanceAmt)
                      inputSheet.getRange(4,1).setValue(numDealFunded);
                      setCounterValuesToSheet();
                      dataRange = inputSheet.getDataRange();
                      dataRangeValues = dataRange.getValues();
                      break;
                    } else if(chkHTisNeg(element,dataRangeValues[i], i, dataRangeValues) === true) {
                      inputSheet.getRange(i+1,inputSheet.getLastColumn()).setValue('Skip H/T').setHorizontalAlignment('normal');
                      dataRange = inputSheet.getDataRange();
                      dataRangeValues = dataRange.getValues();
                    } else if(chkHTisNeg(element,dataRangeValues[i], i, dataRangeValues) === 'Fund H/T') {
                      let lastColInt = dataRangeValues[i].length;
                      dataRangeValues[i][lastColInt-1] = 'Skip H/T';
                      let skipHTarr = dataRangeValues.filter(row => row[lastColInt-1]==='Skip H/T');
                      skipHTarr.forEach(item => {
                        if(item[5]<=balanceAmt){
                          if(regionCounter.get(element)>0) {
                            chkHTAndSetCounters(item);
                            inputSheet.createTextFinder(item[0]).findNext().offset(0,lastColInt-1).setValue('Fund skpdH/T').setHorizontalAlignment('normal').getDataRegion(SpreadsheetApp.Dimension.COLUMNS).setBackground('#9bbb59');
                            ss.getSheetByName('Funded Projects').appendRow(item);
                            balanceAmt -= item[5];
                            numDealFunded++
                          } else {
                            inputSheet.createTextFinder(item[0]).findNext().offset(0,lastColInt-1).setValue('Skip Negative').setHorizontalAlignment('normal');
                          }
                        } else {
                          inputSheet.createTextFinder(item[0]).findNext().offset(0,lastColInt-1).setValue('Skip H/T').setHorizontalAlignment('normal');
                        }
                      });
                      inputSheet.getRange(3,1).setValue(balanceAmt)
                      inputSheet.getRange(4,1).setValue(numDealFunded);
                      setCounterValuesToSheet();
                      dataRange = inputSheet.getDataRange();
                      dataRangeValues = dataRange.getValues();
                    }
                  }else {
                    inputSheet.getRange(i+1,inputSheet.getLastColumn()).setValue('Skip Negative').setHorizontalAlignment('normal');
                    dataRange = inputSheet.getDataRange();
                    dataRangeValues = dataRange.getValues();
                    break;
                  }
                }else{
                  inputSheet.getRange(i+1,inputSheet.getLastColumn()).setValue('Skip 125').setHorizontalAlignment('normal');
                  dataRange = inputSheet.getDataRange();
                  dataRangeValues = dataRange.getValues();

                }
              }
            }
          }
        }
      }
    }
    runCycle++;
  }
};

/**Function to check if any other HT exists with equal point score before funding HT where counter in housing type goes negative. Mark as Skip/HT if so.*/
function chkHTisNeg(sheetName, currData, rowNum, dataSet){
  let inputSheet = ss.getSheetByName(sheetName);
  let i = rowNum;
  if(currData[2] === 'Large Family' && currData[14] === 'Yes'){
    if(hsgTypCounter.get('Large Family High Opportunity')<=0){
      let j=rowNum+1;
      while(j<=dataSet.length){
        if(j===dataSet.length){
          return false; //False here means the row being evaluated is the last one in the data set and hence can be funded.
        } else if(dataSet[j][6] >= currData[6] && dataSet[j][2] !== 'Large Family' && dataSet[j][14] !== 'Yes'){
          return true; //True menas HT has gone negative and the row is not the last one in the data set.
        } 
        j++;
      }
    } else{
      return false; // There is more than a dollar left in the HT and it is not negative. Project can be funded.
    }
  } else {
    for (const bit of hsgTypCounter.entries()){
      if(bit[0] === currData[2]){
        if(bit[1]<=0){
          let j=rowNum+1;
          while(j<=dataSet.length){
            if(j===dataSet.length){
              return 'Fund H/T'; //Fund H/T here means the row being evaluated is the last one in the data set and hence can be funded.
            } else if(dataSet[j][6] >= currData[6] && dataSet[j][2] !== bit[0]){
              return true; //True menas HT has gone negative and the row is not the last one in the data set.
            }
            j++;
          }
        } else {
          return false; // There is more than a dollar left in the HT and it is not negative. Project can be funded.
        }
      }
    }
  }
};

/**Function to check the Housing type in Regions Ta  and deduct the amount from Credit Counter Map if function chkHtisNeg does not return true;*/
function chkHTAndSetCounters(dataRangeValue){
  if(dataRangeValue[2] === 'Large Family' && dataRangeValue[14] === 'Yes'){
    setCounters('hsgTyp','Large Family High Opportunity',dataRangeValue[5]);
    setCounters('hsgTyp','Large Family',dataRangeValue[5]);
    setCounters('region',dataRangeValue[9],dataRangeValue[5]);
  } else {
    setCounters('hsgTyp',dataRangeValue[2],dataRangeValue[5]);
    setCounters('region',dataRangeValue[9],dataRangeValue[5]);
  }
};

/** Function to set counters in the Counter Maps.
 * Counter Names: setAside, hsgTyp, ruralHsgTyp, region; 
 * typName: 
 *  SetAside Types: Nonprofit, RHS & HOME Apportionment, Native American, Other, At-Risk, Special Needs
 *  Housing Types: Large Family, Large Family High Opportunity, Special Needs, At-Risk, Seniors
 *  Rural Housing Types: Acquisition/Rehabilitation, Large Family High Opportunity, Seniors
 *  Geographic Types: City of Los Angeles, Balance of Los Angeles County, Central Valley Region, San Diego County, Inland Empire Region, East Bay Region, Orange County Region, South and West Bay Region, Capital Region, Central Coast Region, Northern Region, San Francisco County*/
function setCounters(counterName, typName, amntDecrIncr){
  let x = amntDecrIncr;
  switch(counterName){
    case 'setAside':
      sACounter.set(typName,sACounter.get(typName)-x);
      break;
    case 'hsgTyp':
      hsgTypCounter.set(typName,hsgTypCounter.get(typName)-x);
      break;
    case 'ruralHsgTyp':
      ruralCounter.set(typName,ruralCounter.get(typName)-x);
      break;
    case 'region':
      regionCounter.set(typName,regionCounter.get(typName)-x);
      break;
  }
};
/** Function to Set the map counters to the Credit Counter Sheet */
function setCounterValuesToSheet(){
  //Set Housing Type
  for(let j=1; j<hTValues.length; j++){
  hTValues[j][2] = hsgTypCounter.get(hTValues[j][1]);
  }
  ss.getRangeByName('HousingType_Counter').setValues(hTValues).setNumberFormat('#,###');
  //Set Set-Aside Type
  for(let j=1; j<sAValues.length; j++){
    sAValues[j][2] = sACounter.get(sAValues[j][1]);
  }
  ss.getRangeByName('SetAside_Counter').setValues(sAValues).setNumberFormat('#,###');
  //Set Rural Housing Type
  for(let j=1; j<ruralHTValues.length; j++){
    ruralHTValues[j][2] = ruralCounter.get(ruralHTValues[j][1]);
  }
  ss.getRangeByName('Rural_Counter').setValues(ruralHTValues).setNumberFormat('#,###');
  //Set Geographic Regions
  for(let j=1; j<regionValues.length; j++){
    regionValues[j][2] = regionCounter.get(regionValues[j][1]);
  }
  ss.getRangeByName('GeoRegion_Counter').setValues(regionValues).setNumberFormat('#,###');
};

/**------------------------------------------------------------------------------------------------------------------**/
function resetCreditCounter(){
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let creditCounterSheet = ss.getSheetByName('CreditCounter');
  let roundNum = ss.getSheetByName('Credit Amount').getRange(4,11).getValue();
  let mapSAsideArr = ss.getRangeByName('Set_Aside_Range').getValues().map(col => [col[7]]).slice(2,10);
  //mapSAsideArr.splice(1,1);
  mapSAsideArr.splice(4,2);
  console.log(mapSAsideArr);
  let setAsideArr = [["='Credit Amount'!K9"],["='Credit Amount'!K11"],["='Credit Amount'!K12"],["='Credit Amount'!K13"],["='Credit Amount'!K15"],["='Credit Amount'!K16"]];
  creditCounterSheet.getRange(2,3,6).setValues(mapSAsideArr);
  let mapHTArray = new Array
  if(roundNum === 'R1'){
    mapHTArray = ss.getRangeByName('Housing_Type_Range').getValues().map(col => [col[5]]).splice(1);
  console.log(mapHTArray);
  } else {
    mapHTArray = ss.getRangeByName('Housing_Type_Range').getValues().map(col => [col[6]]).splice(1);
  console.log(mapHTArray);
  };
  let mapRuralHTArr = new Array;
  if(roundNum === 'R1'){
    mapRuralHTArr = ss.getRangeByName('Rural_Housing_Type_Range').getValues().map(col => [col[5]]).splice(1);
  console.log(mapRuralHTArr);
  } else {
    mapRuralHTArr = ss.getRangeByName('Rural_Housing_Type_Range').getValues().map(col => [col[6]]).splice(1);
  console.log(mapRuralHTArr);
  };
  let mapGeoArr = ss.getRangeByName('Geo_Region_Range').getValues().map(col=>[col[9]]).slice(1,13);
  creditCounterSheet.getRange(9,3,6).setValues(mapHTArray);
  creditCounterSheet.getRange(16,3,3).setValues(mapRuralHTArr);
  creditCounterSheet.getRange(20,3,12).setValues(mapGeoArr);

  /**--Clear Funded Projects Sheet--**/
  const fundSheet = ss.getSheetByName('Funded Projects');
  let dataResetRange = fundSheet.getDataRange().offset(1,0);
  dataResetRange.clear();

  /**--Delete Set-Aside and Geo Sheets--**/
  const sheets = ss.getSheets();
  const numSheets = ss.getNumSheets();
  for(let i=5; i<numSheets; i++){
    ss.deleteSheet(sheets[i]);
  }
};

/**--Repeater Function--*/
function repeat(n, action) {
  for (let i = 0; i < n; i++) {
    action(i);
  }
}

let a = counter.hsgTyp;
let b = Array.from(a.entries())
function x(){
  console.log(b);
};

/**-------------*--------------Code in Abeyance-------------*---------------------------
if(ruralCounter.get('Acquisition/Rehabilitation')>1 && ruralCounter.get('Large Family High Opportunity')>1 && ruralCounter.get('Seniors')>1){
  inputSheet.getRange(i+1,inputSheet.getLastColumn()).setValue('Fund').setHorizontalAlignment('normal');
  ss.getSheetByName('Funded Projects').appendRow(dataRangeValues[i]);
  let x = sACounter.get('RHS & HOME Apportionment');
  sACounter.set('RHS & HOME Apportionment',x-dataRangeValues[i][3]);
  chkRrlHTAndSetCounters(dataRangeValues[i]);
} else if(ruralCounter.get('Large Family High Opportunity')<=1){
  if(dataRangeValues[i][2] === 'Large Family' && dataRangeValues[i][14] === 'Yes'){
    inputSheet.getRange(i+1,inputSheet.getLastColumn()).setValue('Skip H/T').setHorizontalAlignment('normal');
  }
} 
*/


