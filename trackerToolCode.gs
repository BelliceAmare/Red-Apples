/**
 * This program tracks lab minutes based on labs completed.
 * And labs incompleted. Will write a better description l8 er.
 * 
 */

// Creating the spreadsheet object, sheet object and range of the whole sheet.
var ss = SpreadsheetApp.getActiveSpreadsheet();
var trackerSheet = ss.getSheets()[0];
var trackerRange = trackerSheet.getDataRange();
var trackerData = trackerRange.getValues();

var trackerRows = trackerRange.getNumRows();
var trackerCols = trackerRange.getNumColumns();


// Setting named ranged for easier access of data
ss.setNamedRange("labs", ss.getRange("F2:J8"));
ss.setNamedRange("addMinutes", ss.getRange("E2:E"+trackerRows));
ss.setNamedRange("totalMinutes", ss.getRange("D2:D"+trackerRows));
ss.setNamedRange("emails", ss.getRange("C2:C"+trackerRows));
ss.setNamedRange("firstNames", ss.getRange("A2:A"+trackerRows));
ss.setNamedRange("lastNames",ss.getRange("B2:B"+trackerRows));

// Giving variable names for each range.
var labRange = ss.getRangeByName("labs");
var minsRange = ss.getRangeByName("totalMinutes"); 
var emailRange = ss.getRangeByName("emails");
var earnedRange = ss.getRangeByName("addMinutes");
var firstNameRange = ss.getRangeByName("firstNames");
var lastNameRang = ss.getRangeByName("lastNames");



// Getting named range values into variables for organization.
var labData = ss.getRangeByName("labs").getValues();
var minsData = ss.getRangeByName("totalMinutes").getValues(); 
var emailData = ss.getRangeByName("emails").getValues();
var earnedData = ss.getRangeByName("addMinutes").getValues();
var firstNameData = ss.getRangeByName("firstNames").getValues();
var lastNameData = ss.getRangeByName("lastNames").getValues();
// Lab minutes threshold
var alertThreshold = minsRange.getValues()[0][0];


// Spreadsheet functions
function trackerSheetCheck(){
  formatingCheck();
  minutesCheck();
  minutesThreshold();
}
function emailAlerts(){

}

// Data functions (to be passed to spreadShett functions)
function passOrFail(){
  var failCon = SpreadsheetApp.newConditionalFormatRule()
  .whenNumberEqualTo(1)
  .setBackground("#FF0233")
  .setRanges([ss.getRange("labs")])
  .build();

  var passCon = SpreadsheetApp.newConditionalFormatRule()
  .whenNumberEqualTo(0)
  .setBackground("#051CFC")
  .setRanges([ss.getRange("labs")])
  .build();

  var labRules = trackerSheet.getConditionalFormatRules();
  labRules.push(failCon, passCon);
  trackerSheet.setConditionalFormatRules(labRules);
}
function formatingCheck(){
  var isCondition = trackerSheet.getConditionalFormatRules();
  if (isCondition.length == 0){
    passOrFail();
  }
  else{
  // Checking to see how many conditional formatting inthis sheet.
    console.log("This Sheet has " + isCondition.length + " conditional rules");

  // Remove any other condition if is it not passOrFail condition
  // Then apply pass or fail condition
      
    for (var i in isCondition){
      var criteriaVal = isCondition[i].getBooleanCondition().getCriteriaValues();
      if (criteriaVal[i] != 1 || 0){
        trackerSheet.clearFormats();
        passOrFail();
        }
      }
  }
}
function minutesCheck(){
  var addedMinutes = [];
  // Getting the number of passes labs from the labData range
  labData.forEach(function(data){
    var value = data.filter(x=> x==1).length;
    addedMinutes.push(value);
  });

  addedMinutes = (addedMinutes.map(x=> [x * 45]));
  earnedRange.setValues(addedMinutes);
}
function minutesThreshold(){
  // Lab Minutes values update
  minsRange.setValue((labRange.getNumColumns())*45);
}
function alertEmails(){
  var emailIndex = [];
  var emailList = [];
  earnedData.forEach(function(x){
    if (x < alertThreshold){
      emailIndex.push(earnedData.indexOf(x)); 
    }
  });
  emailIndex.forEach(function(x){
    email = emailData[x];
    emailList.push(email);
  });
  return (emailList);

}
function sendAlertEmails(){
  var recipient = [].concat.apply([],alertEmails());
  var labminutes = [].concat.apply([],earnedData);
 
  var firstname = [].concat.apply([],firstNameData);
  var lastname = [].concat.apply([],lastNameData);
  
  recipient.forEach(function(x){
    var message = `Hello ${firstname[recipient.indexOf(x)]} ${lastname[recipient.indexOf(x)]},
    \nYou are recieving this email to alert you that you only have ${labminutes[recipient.indexOf(x)]} minutes
    out of the ${alertThreshold}.\n Make time to come in to complete these lab Thursday's after school or something.\n BHSVA
    Science Deparment`;
    console.log(message);
  });
  

}
