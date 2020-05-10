//***********************************************************************************//
//  Helper Function to Check if a Box is really a box RETURNS= TRUE if is a Box
//***********************************************************************************//
function isBox(job_name){

if (job_name == job_name.toUpperCase()){
  return true;
}
else{
 return false;
}

}
//***********************************************************************************//
//
// Deletes/ checks all Boxes
//
//***********************************************************************************//

function deleteBoxes() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var lastRow = sheet.getLastRow();
  var row=1;
 
  
  while (row < lastRow){
  
  var cell = sheet.getRange("A"+row).getValue();
  var boxResult = isBox(cell);
  
  if (cell == ""){break}
  
  else if(boxResult == true){
    sheet.deleteRow(row); 
    
  }
  else{
   row++;
  }
  
  }

}
//***********************************************************************************//
//
//  This function will delete all unnecessary information from the job name and ONLY
//  filter the job name
//***********************************************************************************//

function RetrieveJobs(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  var lastRow = sheet.getLastRow();
  
  for (var row=1; row <= lastRow; row++){
  
    var currentRow =sheet.getRange("A"+row);
    var date = sheet.getRange("A"+row).getValue();
    var dateExpression = date.search(/\d{2}(\D)\d{2}\1\d{4}/g);
   
    //Keep Beginning to dateExpression
    var resultJob = date.substring(0, dateExpression-1);

    if (dateExpression == -1.0){break};
    currentRow.setValue(resultJob);
    
  }
}
//***********************************************************************************//
//
//  Inserting vLook Up formula to fetch from dicitonary
//
//***********************************************************************************//

function insertLookUp(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  var lastRow = sheet.getLastRow();
  
  for (var row=1; row <= lastRow; row++){
  
    var index="=INDEX(Job_Dictionary!A:A,C"+ row +",0)" ;
    var match="=MATCH(A"+ row +",Job_Dictionary!A:A, 0)+1" ;
    
    var currentRow =sheet.getRange("B"+row);
    currentRow.setValue(index);
    
    var currentRow =sheet.getRange("C"+row);
    currentRow.setValue(match);
 
  }

}
//***********************************************************************************//
//
//
//
//***********************************************************************************//

function sortIt(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  sheet.sort(2, false);
  
}
//***********************************************************************************//
//
//  A Test function to check regex of job_test function()
//
//***********************************************************************************//

function runtest(){

var job_test="US_PRD_nET_ORDER_BOX                                             10/21/2019 10:32:34  -----                RU 334114690/1     "
var result = isBox(job_test);
Logger.log(result);

}

//***********************************************************************************//
//
//
//
//***********************************************************************************//

function getContactList(){

  deleteBoxes();
  RetrieveJobs();
  insertLookUp();
  sortIt();
  
}

