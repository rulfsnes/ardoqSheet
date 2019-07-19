//getArdoqReport
function getArdoqGraphReport(orgURL, reportID, token) {
  
  var url = orgURL + '/api/graph-search/' + reportID + '/run'; 
  var options = {
   muteHttpExceptions: true,
   method : 'get',
   headers: {  
     'Accept': 'application/json',
     'Authorization': 'Token ' + 'token=' + token, // Access token
   }
  };
  Logger.log(options)
  var response = UrlFetchApp.fetch(url,options);
  return response;
}


//Converts Ardoq Gremlin JSON to array (Required to perform sheet range update)
function convertArdoqJsonToArray(ardoqResult){
  var sheetArr = []
  ardoqResult.forEach(
    function(items){
      var row = []
      for(var col in items){
        var cell = items[col]
        row.push(cell)
      }
      sheetArr.push(row)
    }
  )
  return sheetArr
}

//Not in use
function flatten(data) {
    var result = {};
    function recurse (cur, prop) {
        if (Object(cur) !== cur) {
            result[prop] = cur;
        } else if (Array.isArray(cur)) {
             for(var i=0, l=cur.length; i<l; i++)
                 recurse(cur[i], prop + "[" + i + "]");
            if (l == 0)
                result[prop] = [];
        } else {
            var isEmpty = true;
            for (var p in cur) {
                isEmpty = false;
                recurse(cur[p], prop ? prop+"."+p : p);
            }
            if (isEmpty && prop)
                result[prop] = {};
        }
    }
    recurse(data, "");
    return result;
}

//Fetches Ardoq environment settings from 'Ardoq-Settings' sheet in document. Required!!
function getArdoqSettings(){
  var sheetSettings = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Ardoq-Settings')
  var targetSheet = onSearch2('targetSheet', sheetSettings)
  var reportID = onSearch2('reportID', sheetSettings)
  var orgURL = onSearch2('orgURL', sheetSettings)
  var token = onSearch2('token', sheetSettings)
  return {targetSheet: targetSheet, reportID: reportID, orgURL : orgURL, flatten: flatten}
}


//Main function to update sheet
function updateArdoqReport(){
  ardoqSettings = getArdoqSettings()
  var response = getArdoqGraphReport(ardoqSettings.orgURL, ardoqSettings.reportID,'138073e5eecd4a11a66f86e169fae327');

  var reportObj = JSON.parse(response);
  var result = reportObj.result
  var firstRow = Object.keys(result[0]) //Uses the first result to define header row
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ardoqSettings.targetSheet);
  // write result
  
  var Arr = convertArdoqJsonToArray(result);
 
  sheet.getRange(1,1,1,firstRow.length).setValues([firstRow])
  sheet.getRange(2,1,Arr.length,Arr[0].length).setValues(Arr);

}


//function onOpen provides button to run report in Google Sheet
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Update Ardoq Report')
      .addItem('Update report','updateArdoqReport')
      .addToUi();
}

//Searches for value in google sheet
function onSearch2(searchString, sheet) {
  var values = sheet.getDataRange().getValues();
  var i = findIndex(values, searchString);
  
  return values[i][1]; 
}

//Finds index of given searchString in row. 
function findIndex(values, searchString) {
  for(var i=0, iLen=values.length; i<iLen; i++) {
    if(values[i][0] == searchString) {
      return i;
    }
  }
}