function duplicateSheet() {
  SpreadsheetApp.getActiveSpreadsheet().duplicateActiveSheet();
  let activeSS = SpreadsheetApp.getActiveSheet();
  activeSS.setName('Sort By Type');
}

/**------------------------------------------------------------------------------------------------------------------**/
function nonProfitSort(){
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let creditDataSheet = ss.getSheets()[0];
  let rawDataSheet = ss.getSheets()[2];
  let setAsideCreditSheetValues = ss.getRangeByName('Set_Aside_Range').getValues();//Get's Values From Set_Aside_Range in 'Credit Amount' Sheet
  ss.insertSheet('NonProfit',3);

  //create 'NonProfit' sheet
  let nonProfitSheet = ss.getSheetByName('NonProfit');
  ss.setActiveSheet(rawDataSheet);
  
  //Set Credit Amount and Header Row in NonProfit Sheet
  let rawDataSheetHeaderRange = rawDataSheet.getRange(1,1,1,rawDataSheet.getLastColumn());
  let rawDataSheetHeaderValues = rawDataSheetHeaderRange.getValues();
  nonProfitSheet.getRange(1,1).setValue(setAsideCreditSheetValues[2][7]).setNumberFormat('#,###');
  nonProfitSheet.getRange(1,2).setValue(setAsideCreditSheetValues[2][0]);
  nonProfitSheet.getRange(2,1,rawDataSheetHeaderRange.getHeight(),rawDataSheetHeaderRange.getWidth()).setValues(rawDataSheetHeaderValues).setFontWeight('bold').setHorizontalAlignment('center').setWrap(true);

  //Get Data for Projects qualifying under NonProfit Set-Aside
  let rawDataRange = rawDataSheet.getDataRange();
  let rawDataValues = rawDataRange.getValues().slice(1);
  let nonProfitArr = new Array;
  for(let i=0; i<rawDataValues.length; i++){
    let setAsidei7Array = rawDataValues[i][7].split(' ')
    let setAsidei7FirstValue = setAsidei7Array[0];
    if(setAsidei7FirstValue == 'Nonprofit'){
      console.log(setAsidei7FirstValue);
      nonProfitArr.push(rawDataValues[i]);
    }
  }

  //Set and Formatting for Projects qualifying under NonProfit Set-Aside
  nonProfitSheet.getRange(3,1,nonProfitArr.length,nonProfitArr[0].length).setValues(nonProfitArr).setFontWeight('normal').setHorizontalAlignment('right').setWrap(false).setNumberFormat('#,###');
  nonProfitSheet.setColumnWidths(1,9,100);
  nonProfitSheet.autoResizeColumn(2).autoResizeColumn(8).autoResizeColumns(9,4).hideColumns(13,nonProfitSheet.getLastColumn()-13+1);
  nonProfitSheet.hideColumns(10,2);
  nonProfitSheet.getRange('G:G').setNumberFormat('##.##%');

  //Sort Data by Points then homeless and then Tie-breaker
  let nonProfitDataRange = nonProfitSheet.getRange(3,1,nonProfitSheet.getLastRow(),nonProfitSheet.getLastColumn());
  nonProfitDataRange.sort([{column: 6, ascending: false},{column: 8, ascending: true}, {column: 7, ascending: false}]);

  //Create Combined Credit Requested Column
  nonProfitSheet.insertColumnAfter(5);
  nonProfitSheet.getRange('F2').setValue('COMBINED');
  nonProfitSheet.getRange(3,6,nonProfitSheet.getLastRow()-3+1,1).setFormulaR1C1("=((R[0]C[-2]*10)+R[0]C[-1])/10");
  ss.setActiveSheet(nonProfitSheet);

  //Funding code
  
  let x = nonProfitSheet.getRange(1,1).getValue();
  let hTValues = ss.getRangeByName('HousingType_Counter').getValues();
  let nonProfitSheetFundRange = nonProfitSheet.getRange(3,1,nonProfitSheet.getLastRow()-3,nonProfitSheet.getLastColumn()+1);
  let nonProfitSheetFundValues = nonProfitSheetFundRange.getValues();
  let nonProfitSheetFundColumnNumber = nonProfitSheet.getLastColumn()+1;
  for (let i=0; i<nonProfitSheetFundValues.length; i++){
    console.log(x);
    if (x >= 1){
      nonProfitSheet.getRange(i+3,nonProfitSheetFundColumnNumber).setValue('Fund');
      x = x - nonProfitSheetFundValues[i][3];
      console.log(x);
      for(let j=0; j<hTValues.length; j++){
        console.log(hTValues[j][2]);
        if(hTValues[j][1] === nonProfitSheetFundValues[i][2] && nonProfitSheetFundValues[i][14] === 'Yes'){
          console.log(hTValues[j][2]);
          hTValues[j][2] = hTValues[j][2] - nonProfitSheetFundValues[i][5];
          console.log(hTValues[j][2]);
        } else if (hTValues[j][1] === nonProfitSheetFundValues[i][2]) {
            console.log(hTValues[j][2]);
            hTValues[j][2] = hTValues[j][2] - nonProfitSheetFundValues[i][5];
            console.log(hTValues[j][2]);
        }
      }
    }
  }
  ss.getRangeByName('HousingType_Counter').setValues(hTValues).setNumberFormat('#,###');
  ss.getSheetByName('CreditCounter').getRange(2,3).setValue(x).setNumberFormat('#,###');
  console.log(nonProfitSheetFundValues);
}

/**------------------------------------------------------------------------------------------------------------------**/
function rawDataWrapper(){
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let rawDataSheet = ss.getSheets()[2];
  let rawDataRange = rawDataSheet.getDataRange();
  let rawDataRangeValues = rawDataRange.getValues();
  return rawDataRangeValues;
}

/**------------------------------------------------------------------------------------------------------------------**/
function atRiskSort(){
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheetName = 'At-Risk';
  ss.insertSheet(sheetName,4);
  let atRiskSheet = ss.getSheetByName(sheetName);
  let rawDataSheetValues = rawDataWrapper();
  //console.log(rawDataSheetValues[0]);
  let setAsideCreditSheetValues = ss.getRangeByName('Set_Aside_Range').getValues();//Get's Values From Set_Aside_Range in 'Credit Amount' Sheet
  for(i=2; i<setAsideCreditSheetValues.length; i++){
    if(setAsideCreditSheetValues[i][0] === sheetName){
      atRiskSheet.getRange(1,1).setValue(setAsideCreditSheetValues[i][7]).setNumberFormat('#,###');
      atRiskSheet.getRange(1,2).setValue(setAsideCreditSheetValues[i][0]);
      atRiskSheet.getRange(2,1,1,rawDataSheetValues[0].length).setValues(rawDataSheetValues.slice(0,1)).setFontWeight('bold').setHorizontalAlignment('center').setWrap(true);
    }
  }

  //Get Data for Projects qualifying under At-Risk Set-Aside
  let rawDataValues = rawDataSheetValues.slice(1);
  let atRiskArr = new Array;
  for(let i=0; i<rawDataValues.length; i++){
    let setAsidei7Array = rawDataValues[i][7].split(' ')
    let setAsidei7FirstValue = setAsidei7Array[0];
    if(setAsidei7FirstValue == 'At-Risk'){
      console.log(setAsidei7FirstValue);
      atRiskArr.push(rawDataValues[i]);
    }
  }

  //Set and Formatting for Projects qualifying under At-Risk Set-Aside
  atRiskSheet.getRange(3,1,atRiskArr.length,atRiskArr[0].length).setValues(atRiskArr).setFontWeight('normal').setHorizontalAlignment('right').setWrap(false).setNumberFormat('#,###');
  atRiskSheet.setColumnWidths(1,9,100);
  atRiskSheet.autoResizeColumn(2).autoResizeColumn(8).autoResizeColumns(9,4).hideColumns(13,atRiskSheet.getLastColumn()-13+1);
  atRiskSheet.hideColumns(10,2);
  atRiskSheet.getRange('G:G').setNumberFormat('##.##%');

    //Create Combined Credit Requested Column
  atRiskSheet.insertColumnAfter(5);
  atRiskSheet.getRange('F2').setValue('COMBINED');
  atRiskSheet.getRange(3,6,atRiskSheet.getLastRow()-3+1,1).setFormulaR1C1("=((R[0]C[-2]*10)+R[0]C[-1])/10");
  ss.setActiveSheet(atRiskSheet);

    //Funding code
  let x = atRiskSheet.getRange(1,1).getValue();
  let hTValues = ss.getRangeByName('HousingType_Counter').getValues();
  let atRiskSheetFundRange = atRiskSheet.getRange(3,1,atRiskSheet.getLastRow()-2,atRiskSheet.getLastColumn()+1);
  let atRiskSheetFundValues = atRiskSheetFundRange.getValues();
  let atRiskSheetFundColumnNumber = atRiskSheet.getLastColumn()+1;
  for (let i=0; i<atRiskSheetFundValues.length; i++){
    console.log(x);
    if (x >= 1){
      atRiskSheet.getRange(i+3,atRiskSheetFundColumnNumber).setValue('Fund');
      x = x - atRiskSheetFundValues[i][3];
      console.log(x);
      for(let j=0; j<hTValues.length; j++){
        console.log(hTValues[j][2]);
        if(hTValues[j][1] === atRiskSheetFundValues[i][2] && atRiskSheetFundValues[i][14] === 'Yes'){
          console.log(hTValues[j][2]);
          hTValues[j][2] = hTValues[j][2] - atRiskSheetFundValues[i][5];
          console.log(hTValues[j][2]);
        } else if (hTValues[j][1] === atRiskSheetFundValues[i][2]) {
            console.log(hTValues[j][2]);
            hTValues[j][2] = hTValues[j][2] - atRiskSheetFundValues[i][5];
            console.log(hTValues[j][2]);
        }
      }
    }
  }
  ss.getRangeByName('HousingType_Counter').setValues(hTValues).setNumberFormat('#,###');
  ss.getSheetByName('CreditCounter').getRange(6,3).setValue(x).setNumberFormat('#,###');
  console.log(atRiskSheetFundValues);
}


/**------------------------------------------------------------------------------------------------------------------**/
function setAsideSort(){
  let arrSANames = ['Nonprofit', 'At-Risk', 'Special Needs'];
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let setAsideCreditSheetValues = ss.getRangeByName('Set_Aside_Range').getValues();
  for (const element of arrSANames){
    let regEx = new RegExp(element, 'i');
    for(let i=0; i< setAsideCreditSheetValues.length; i++){
      if(regEx.test(setAsideCreditSheetValues[i][0])){
        let sheetName = element;
        ss.insertSheet(sheetName,ss.getNumSheets()+1);
        let inputSheet = ss.getSheetByName(sheetName);
        let rawDataSheetValues = rawDataWrapper();
        inputSheet.getRange(1,1).setValue(setAsideCreditSheetValues[i][7]).setNumberFormat('#,###');
        inputSheet.getRange(1,2).setValue(setAsideCreditSheetValues[i][0]);
        inputSheet.getRange(2,1,1,rawDataSheetValues[0].length).setValues(rawDataSheetValues.slice(0,1)).setFontWeight('bold').setHorizontalAlignment('center').setWrap(true);

        //Get Data for Projects qualifying under Set-Aside
        let rawDataValues = rawDataSheetValues.slice(1);
        let arrRawDataToInput = new Array;
        for(let j=0; j<rawDataValues.length; j++){
          if(regEx.test(rawDataValues[j][7])){
            arrRawDataToInput.push(rawDataValues[j]);
          } 
        };

        //Set Data and Formatting for Projects qualifying under Set-Aside
        inputSheet.getRange(3,1,arrRawDataToInput.length,arrRawDataToInput[0].length).setValues(arrRawDataToInput).setFontWeight('normal').setHorizontalAlignment('right').setWrap(false).setNumberFormat('#,###');  

        //Create Combined Credit Requested Column
        inputSheet.insertColumnAfter(5);
        inputSheet.getRange('F2').setValue('COMBINED');
        inputSheet.getRange(3,6,inputSheet.getLastRow()-3+1,1).setFormulaR1C1("=((R[0]C[-2]*10)+R[0]C[-1])/10");
        ss.setActiveSheet(inputSheet);

           
        if(element === 'Special Needs'){
          let regExHomeless = new RegExp('homeless','i');
          let nonProfitSheet = ss.getSheetByName('Nonprofit');
          let homelessArr = new Array;
          let nonProfitDataValues = nonProfitSheet.getRange(3,1,nonProfitSheet.getLastRow()-2,nonProfitSheet.getLastColumn()).getValues();
          for(let z=0; z<nonProfitDataValues.length; z++){
            if(regExHomeless.test(nonProfitDataValues[z][8]) && nonProfitDataValues[z][26] !=='Fund'){
              homelessArr.push(nonProfitDataValues[z]);
            }
          }
          inputSheet.getRange(inputSheet.getLastRow()+1,1,homelessArr.length,homelessArr[0].length).setValues(homelessArr).setFontWeight('normal').setHorizontalAlignment('right').setWrap(false).setNumberFormat('#,###');
        };

         //Sort Data by Points then homeless and then Tie-breaker for "Nonprofit" element; ort Data by Points then by Tiebreaker for all else.
        let inputSheetDataRange = inputSheet.getRange(3,1,inputSheet.getLastRow(),inputSheet.getLastColumn());
        if(element === 'Nonprofit'){
          inputSheetDataRange.sort([{column: 7, ascending: false},{column: 9, ascending: true}, {column: 8, ascending: false}]);
        } else {
          inputSheetDataRange.sort([{column: 7, ascending: false}, {column: 8, ascending: false}]);
        };

        //Set Formatting
        inputSheet.setColumnWidths(1,9,100);
        inputSheet.autoResizeColumn(2).autoResizeColumns(9,4).hideColumns(14,inputSheet.getLastColumn()-14+1);
        inputSheet.hideColumns(11,2);
        inputSheet.getRange('H:H').setNumberFormat('##.##%');   

        //Funding code
        let x = inputSheet.getRange(1,1).getValue();
        let hTValues = ss.getRangeByName('HousingType_Counter').getValues();
        let inputSheetFundRange = inputSheet.getRange(3,1,inputSheet.getLastRow()-3+1,inputSheet.getLastColumn()+1);
        let inputSheetFundValues = inputSheetFundRange.getValues();
        let inputSheetFundColumnNumber = inputSheet.getLastColumn()+1;
        for (let k=0; k<inputSheetFundValues.length; k++){
          console.log(x);
          if (x >= 1){
            inputSheet.getRange(k+3,inputSheetFundColumnNumber).setValue('Fund');
            x = x - inputSheetFundValues[k][3];
            console.log(x);
            for(let m=0; m<hTValues.length; m++){
              console.log(hTValues[m][2]);
              if(hTValues[m][1] === inputSheetFundValues[k][2] && inputSheetFundValues[k][14] === 'Yes'){
                console.log(hTValues[m][2]);
                hTValues[m][2] = hTValues[m][2] - inputSheetFundValues[k][5];
                console.log(hTValues[m][2]);
              } else if (hTValues[m][1] === inputSheetFundValues[k][2]) {
                  console.log(hTValues[m][2]);
                  hTValues[m][2] = hTValues[m][2] - inputSheetFundValues[k][5];
                  console.log(hTValues[m][2]);
              }
            }
          }
        }

        //Set and Format Set-Aside Counter and Housing Type Counter
        ss.getRangeByName('HousingType_Counter').setValues(hTValues).setNumberFormat('#,###');
        let setAsideCounterValues = ss.getRangeByName('SetAside_Counter').getValues();
        for(let n=0; n<setAsideCounterValues.length; n++){
          if(setAsideCounterValues[n][1] === element){
            ss.getSheetByName('CreditCounter').getRange(n+1,3).setValue(x).setNumberFormat('#,###');
          }
        }
      }
    }
  }
  
}
/**------------------------------------------------------------------------------------------------------------------**/
function resetCreditCounter(){
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let creditCounterSheet = ss.getSheetByName('CreditCounter');
  let roundNum = ss.getSheetByName('Credit Amount').getRange(4,11).getValue();
  let mapSAsideArr = ss.getRangeByName('Set_Aside_Range').getValues().map(col => [col[7]]).slice(2,10);
  mapSAsideArr.splice(1,1);
  mapSAsideArr.splice(4,1);
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
  creditCounterSheet.getRange(9,3,5).setValues(mapHTArray);
  creditCounterSheet.getRange(15,3,3).setValues(mapRuralHTArr);
};



