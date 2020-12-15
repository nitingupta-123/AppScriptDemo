//1. Spreadsheet
//2. Sheet Inside Spreadsheet. 
//3. NamedRange inside Sheet 
//4. CRUD on data in a namedRange and Sheet

//function doGet(request) {
//  const sh1 = SpreadsheetApp.getActiveSpreadsheet();
//  const sheet = sh1.getSheetByName("Sheet2").getRange("Sheet2!A1:H6").getValues();
//  Logger.log(JSON.stringify(sheet));
//  return ContentService.createTextOutput(JSON.stringify(sheet))
//    .setMimeType(ContentService.MimeType.JSON);
//}

//function doGet(e) {
//   Logger.log(e);
//  Logger.log(JSON.stringify(e));
////  Logger.log(e.parameter['message']);
//  return HtmlService.createHtmlOutputFromFile('Example1');  
//}

//function doGet(e) 
//{
//  Logger.log(JSON.stringify(e));
//  Logger.log(e.parameter['message']);
//  var htmlOutput =  HtmlService.createTemplateFromFile('Example2');
//  htmlOutput.message = e.parameter['message'];
//  Logger.log(htmlOutput.evaluate());
//  return htmlOutput.evaluate();
//}

function doGet(e) 
{
  Logger.log(JSON.stringify(e));
  let htmlOutput =  HtmlService.createTemplateFromFile('Example3');
  if(!e.parameter['spreadsheet']){
  
    htmlOutput.spreadsheet = '';
    
  }else
  {
    // taking Spreadsheet Name and sheet Name and NameRange from html form
    const newSpreadsheetName=e.parameter['spreadsheet'];
    const newSheetName=e.parameter['sheetName'];
    const nameRange=e.parameter['namedRange'];
    
    // creating Spreadsheet
    
    const newSpreadsheetId= createNewSpreadsheet(newSpreadsheetName);
    
    // creating New sheet
    
    createNewSheet(newSpreadsheetId,newSheetName);
    
    // Creating Name range
    let rangeCheckName=createNamedRange(newSpreadsheetId,newSheetName,nameRange);
    
    htmlOutput.spreadsheet = 'Spreadsheet is: ' + newSpreadsheetName+ ' is created  and sheet name is: ' +
    newSheetName+' is created '+ rangeCheckName +'range Name created';
  }
  htmlOutput.url = getUrl();  
  return htmlOutput.evaluate();
}

function getUrl() {
 let url = ScriptApp.getService().getUrl();
 return url;
}


/*
*creating resources
*/

// create new spreadsheet

function createNewSpreadsheet(title) {
    let sheet = Sheets.newSpreadsheet();
    sheet.properties = Sheets.newSpreadsheetProperties();
    sheet.properties.title = title;
    let spreadsheet = Sheets.Spreadsheets.create(sheet);
//  taking id of newly created spreadsheet
    const  id=spreadsheet.spreadsheetId;
    return id;
}

// create new sheet

function createNewSheet(id,sheetName){
  const activeSpreadsheet = SpreadsheetApp.openById(id)
  // creating new sheet of given name
  const NewSheet = activeSpreadsheet.insertSheet();
  //giving name to newly created sheet
  NewSheet.setName(sheetName);
  const msg=deleteSheet(id,"Sheet1");
  Logger.log(msg);
  return id;
}


// create name range

function createNamedRange(id,sheetName,nameRange) {
  nameRange=nameRange.replace(/ /g,'_');
  Logger.log(nameRange);
  let ss = SpreadsheetApp.openById(id);
  let range = ss.getRange(sheetName+'!A1:E11');
  ss.setNamedRange(nameRange,range);
  let rangeCheck = ss.getRangeByName(nameRange);
  let rangeCheckName = rangeCheck.getA1Notation();
  Logger.log(rangeCheckName);
  return rangeCheckName;
}


/*
*deleting resources
*/

function deleteSheet(id,sheetName){
  var ss = SpreadsheetApp.openById(id);
  var sheet = ss.getSheetByName(sheetName);
  Logger.log(sheet);
  let messsage="";
  if(sheet){
    ss.deleteSheet(sheet);
    message=sheetName+" is deleted successfully!";
  }else{
    message="sheet with name "+sheetName+" does not exist!"
  }
  return message;
}


function deleteNameRange(id,nameRange){
  // The code below deletes all the named ranges in the spreadsheet.
  var namedRanges = SpreadsheetApp.openById(id).getNamedRanges();
  for (var i = 0; i < namedRanges.length; i++) {
    if(namedRanges[i]===namedRange){
      namedRanges[i].remove();
      let message=namedRanges +" is removed sucessfully!";
      return message;
    }
  }
  return "Probably the "+nameRange+" do not exist!";
}


